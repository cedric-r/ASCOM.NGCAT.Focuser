/**
 * ASCOM.NGCAT.Focuser - Robofocus controller
 * Copyright (C) 2021 Cedric Raguenaud [cedric@raguenaud.earth]
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 *
 */
using ASCOM.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ASCOM.NGCAT
{
    /// <summary>
    /// The resources shared by all drivers and devices, in this example it's a serial port with a shared SendMessage method
    /// an idea for locking the message and handling connecting is given.
    /// In reality extensive changes will probably be needed.
    /// Multiple drivers means that several applications connect to the same hardware device, aka a hub.
    /// Multiple devices means that there are more than one instance of the hardware, such as two focusers.
    /// In this case there needs to be multiple instances of the hardware connector, each with it's own connection count.
    /// </summary>
    public static class SharedResources
    {
        // object used for locking to prevent multiple drivers accessing common code at the same time
        private static readonly object lockObject = new object();

        // Shared serial port. This will allow multiple drivers to use one single serial port.
        private static readonly ASCOM.Utilities.Serial ASCOMSerial = new ASCOM.Utilities.Serial();

        // Counter for the number of connections to the serial port
        private static int ConnectedClients = 0;

        private static int CMD_LENGTH = 9;
        // 
        private static string CMD_START { get { return ""; } }
        private static string CMD_END { get { return ""; } }

        public static bool TraceEnabled = false;

        //
        // Public access to shared resources
        //

        /// <summary>
        /// Shared serial port
        /// </summary>
        private static ASCOM.Utilities.Serial SharedSerial
        {
            get
            {
                return ASCOMSerial;
            }
        }

        /// <summary>
        /// number of connections to the shared serial port
        /// </summary>
        public static string COMPortName
        {
            get
            {
                return SharedSerial.PortName;
            }
            set
            {
                if (SharedSerial.Connected && SharedSerial.PortName != value)
                {
                    LogMessage("SharedResources::COMPortName", "NotSupportedException: Serial port already connected");
                    throw new NotSupportedException("Serial port already connected");
                }

                SharedSerial.PortName = value;
                LogMessage("SharedResources::COMPortName", "New serial port name: {0}", value);
            }
        }

        /// <summary>
        /// number of connections to the shared serial port
        /// </summary>
        public static int Connections
        {
            get
            {
                //LogMessage("Connections", "ConnectedClients: {0}", ConnectedClients);
                return ConnectedClients;
            }
            set
            {
                ConnectedClients = value;
                //LogMessage("Connections", "ConnectedClients new value: {0}", ConnectedClients);
            }
        }

        /// <summary>
        /// Bla bla bla wrapper
        /// </summary>
        public static bool Connected
        {
            get
            {
                LogMessage("SharedResources::Connected", "SharedSerial.Connected: {0}", SharedSerial.Connected);
                return SharedSerial.Connected;
            }
            set
            {
                //                if (SharedSerial.Connected == value) { return; }

                // Check if we are the first client using the shared serial
                if (value)
                {
                    LogMessage("SharedResources::Connected", "New connection request");

                    if (Connections == 0)
                    {
                        LogMessage("SharedResources::Connected", "This is the first client");

                        // Check for a valid serial port name
                        if (Array.IndexOf(SharedSerial.AvailableCOMPorts, SharedSerial.PortName) > -1)
                        {
                            lock (lockObject)
                            {
                                // Sets serial parameters
                                SharedSerial.Speed = SerialSpeed.ps9600;
                                SharedSerial.ReceiveTimeout = 5;
                                SharedSerial.Connected = true;

                                Connections++;
                                LogMessage("SharedResources::Connected", "Connected successfully");
                            }
                        }
                        else
                        {
                            LogMessage("SharedResources::Connected", "Connection aborted, invalid serial port name");
                        }
                    }
                    else
                    {
                        lock (lockObject)
                        {
                            Connections++;
                            LogMessage("SharedResources::Connected", "Connected successfully");
                        }
                    }
                }
                else
                {
                    LogMessage("SharedResources::Connected", "Disconnect request");

                    lock (lockObject)
                    {
                        // Check if we are the last client connected
                        if (Connections == 1)
                        {
                            SharedSerial.ClearBuffers();
                            SharedSerial.Connected = false;
                            LogMessage("SharedResources::Connected", "This is the last client, disconnecting the serial port");
                        }
                        else
                        {
                            LogMessage("SharedResources::Connected", "Serial connection kept alive");
                        }

                        Connections--;
                        LogMessage("SharedResources::Connected", "Disconnected successfully");
                    }
                }
            }
        }

        public static char Checksum(string message)
        {
            char val = CalculateChecksum(message.Substring(0, 8)); // Remove the checksum from calculation
            if (message.Length == 9) // Just in case an idiot tries to check the checksum of a non-checksumed message. I've been that idiot.
            {
                if (val != message[8])
                {
                    LogMessage("SharedResources::Checksum", "Bacd checksum: " + message + " should have " + val);
                }
            }

            return val;
        }

        public static char CalculateChecksum(string message)
        {
            message = message.Substring(0, 8);
            int checksum = 0;
            byte[] asciiBytes = Encoding.ASCII.GetBytes(message);
            for (int i = 0; i < asciiBytes.Length; i++)
            {
                checksum += message[i];
            }
            return Convert.ToChar(checksum % 256);
        }

        public static string FormatCommand(string message) // Just the numbers in the command
        {
            if (message.Length == 6)
            {
                // do nothing
            }
            else
            if (message.Length == 5)
            {
                message = "0" + message;
            }
            else
            if (message.Length == 4)
            {
                message = "00" + message;
            }
            else
            if (message.Length == 3)
            {
                message = "000" + message;
            }
            else
            if (message.Length == 2)
            {
                message = "0000" + message;
            }
            else
            if (message.Length == 1)
            {
                message = "00000" + message;
            }
            else
            if (message.Length == 0)
            {
                message = "000000";
            }
            return message;
        }

        public static string ExtractNumber(string message)
        {
            string number = "";
            if (message.Length >= 8 && message[0] == 'F')
            {
                number = message.Substring(2, 6);
            }
            else if (message[0] == 'F')
            {
                number = message.Substring(2);
            }
            else if (message.Length == 6)
            {
                number = message;
            }
            return number;
        }

        public static double ParseNumber(string number)
        {

            double temp = 0;
            try
            {
                LogMessage("SharedResources::ParseNumber", "Parsing " + number);
                if (number.Length >= 8 && number[0] == 'F')
                {
                    string extracted = ExtractNumber(number);
                    LogMessage("SharedResources::ParseNumber", "Number is " + extracted);
                    temp = Double.Parse(Convert.ToDecimal(extracted, System.Globalization.CultureInfo.InvariantCulture).ToString("F5"));
                }
                else if (number.Length == 6)
                {
                    LogMessage("SharedResources::ParseNumber", "Number is " + number);
                    temp = Double.Parse(Convert.ToDecimal(number, System.Globalization.CultureInfo.InvariantCulture).ToString("F5"));
                }
            }
            catch (Exception e)
            {
                LogMessage("SharedResources::ParseNumber", "Error: " + e.Message + "\n" + e.StackTrace);
            }
            return temp;
        }

        public static int ParseNumberAsInt(string number)
        {
            double temp = ParseNumber(number);
            return (int)temp;
        }

        public static double ParseNumberAsDouble(string number)
        {
            return ParseNumber(number);
        }

        public static string ParseCommand(string message)
        {
            string temp = "";
            if (message[0] == 'I') temp = message[0].ToString();
            else
            if (message[0] == 'O') temp = message[0].ToString();
            else
            if (message[0] == 'F') temp = message.Substring(0, 2);
            return temp;
        }

        public static string ReceiveN(int number)
        {
            string temp = "";
            try
            {
                return SharedSerial.ReceiveCounted(number);
            }
            catch (Exception) { }
            return temp;
        }

        public static string Receive()
        {
            return ReceiveN(CMD_LENGTH);
        }

        public static string SendSerialMessage(string message)
        {
            string retval = String.Empty;

            if (SharedSerial.Connected)
            {
                lock (lockObject)
                {
                    SharedSerial.ClearBuffers();
                    message = message.Substring(0, 2) + FormatCommand(ExtractNumber(message));
                    SharedSerial.Transmit(CMD_START + message + CMD_END + CalculateChecksum(message));
                    LogMessage("SharedResources::SendSerialMessage", "Message: {0}", CMD_START + message + CMD_END + CalculateChecksum(message));

                    try
                    {
                        retval = Receive();
                        Checksum(retval); // Check the checksum is right
                        LogMessage("SharedResources::SendSerialMessage", "Message received: {0}", retval);
                    }
                    catch (Exception e)
                    {
                        LogMessage("SharedResources::SendSerialMessage", "Serial timeout exception while receiving data: " + e.Message + "\n" + e.StackTrace);
                    }

                    LogMessage("SharedResources::SendSerialMessage", "Message sent: {0} received: {1}", CMD_START + message + CMD_END, retval);
                }
            }
            else
            {
                //throw new NotConnectedException("SendSerialMessage");
                LogMessage("SharedResources::SendSerialMessage", "NotConnectedException");
            }

            return retval;
        }

        public static void SendSerialMessageBlind(string message)
        {
            if (SharedSerial.Connected)
            {
                lock (lockObject)
                {
                    message = message.Substring(0, 2) + FormatCommand(ExtractNumber(message));
                    SharedSerial.Transmit(CMD_START + message + CMD_END + CalculateChecksum(message));
                    LogMessage("SharedResources::SendSerialMessage", "Message: {0}", CMD_START + message + CMD_END + CalculateChecksum(message));
                }
            }
            else
            {
                //throw new NotConnectedException("SendSerialMessageBlind");
                LogMessage("SharedResources::SendSerialMessageBlind", "NotConnectedException");
            }
        }

        public static double ConvertTemperature(double temperature)
        {
            return (temperature / 2.0) - 273.15;
        }

        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            LogMessage(DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + ": " + identifier + ": " + msg);
        }

        static readonly object fileLockObject = new object();
        internal static void LogMessage(string message)
        {
            lock (fileLockObject)
            {
                if (TraceEnabled) File.AppendAllText(@"c:\temp\Focuser.log", message + "\n");
            }
        }
    }
}
