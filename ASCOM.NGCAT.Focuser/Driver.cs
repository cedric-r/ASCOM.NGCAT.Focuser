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
#define Focuser

using ASCOM;
using ASCOM.Astrometry;
using ASCOM.Astrometry.AstroUtils;
using ASCOM.DeviceInterface;
using ASCOM.Utilities;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace ASCOM.NGCAT
{
    /// <summary>
    /// ASCOM Focuser Driver for NGCAT.
    /// </summary>
    [Guid("835503b2-4cbd-4675-b5ad-40f2af06180a")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Focuser.driverID)]
    public class Focuser : IFocuserV2
    {
        /// <summary>
        /// ASCOM DeviceID (COM ProgID) for this driver.
        /// The DeviceID is used by ASCOM applications to load the driver at runtime.
        /// </summary>
        internal const string driverID = "ASCOM.NGCAT.Focuser";

        /// <summary>
        /// Driver description that displays in the ASCOM Chooser.
        /// </summary>
        private static string driverDescription = "ASCOM NGCAT Driver for Focuser";

        protected static string GET_CURRENT_POS { get { return "FG"; } }
        protected static string GET_TEMPERATURE { get { return "FT"; } }
        protected static string GET_VERSION { get { return "FV"; } }
        protected static string SET_TARGET_POS { get { return "FG"; } }
        protected static string START_ADC_READING { get { return "FV"; } }
        protected static string STOP_MOTION { get { return "FQ"; } }
        protected static string POWER_MODULE { get { return "FP"; } }
        protected static string BACKLASH { get { return "FB"; } }
        protected static string MAXIMUM_TRAVEL { get { return "FL"; } }
        protected static string CONFIGURATION { get { return "FC"; } }
        protected static string SYNC { get { return "FS"; } }

        internal static string comPortProfileName = "COM Port"; // Constants used for Profile persistence
        internal static string comPortDefault = "COM1";
        internal static string traceStateProfileName = "Trace Level";
        internal static string traceStateDefault = "false";

        internal static string comPort; // Variables to hold the currrent device configuration

        /// <summary>
        /// Private variable to hold the connected state
        /// </summary>
        private bool connectedState;
        private bool isMoving = false;
        static readonly object lockObject = new object();

        /// <summary>
        /// Private variable to hold an ASCOM Utilities object
        /// </summary>
        private Util utilities;

        /// <summary>
        /// Private variable to hold an ASCOM AstroUtilities object to provide the Range method
        /// </summary>
        private AstroUtils astroUtilities;

        /// <summary>
        /// Initializes a new instance of the <see cref="NGCAT"/> class.
        /// Must be public for COM registration.
        /// </summary>
        public Focuser()
        {
            ReadProfile(); // Read device configuration from the ASCOM Profile store
            SharedResources.LogMessage("Focuser", "Starting initialisation");

            SharedResources.LogMessage("Focuser", "Port is " + comPort);

            connectedState = false; // Initialise connected to false
            utilities = new Util(); //Initialise util object
            astroUtilities = new AstroUtils(); // Initialise astro utilities object
            //TODO: Implement your additional construction here

            SharedResources.LogMessage("Focuser", "Completed initialisation");
        }


        //
        // PUBLIC COM INTERFACE IFocuserV2 IMPLEMENTATION
        //

        #region Common properties and methods.

        /// <summary>
        /// Displays the Setup Dialog form.
        /// If the user clicks the OK button to dismiss the form, then
        /// the new settings are saved, otherwise the old values are reloaded.
        /// THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
        /// </summary>
        public void SetupDialog()
        {
            // consider only showing the setup dialog if not connected
            // or call a different dialog if connected
            if (IsConnected)
                System.Windows.Forms.MessageBox.Show("Already connected, just press OK");

            using (SetupDialogForm F = new SetupDialogForm())
            {
                var result = F.ShowDialog();
                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    WriteProfile(); // Persist device configuration values to the ASCOM Profile store
                }
            }
        }

        public ArrayList SupportedActions
        {
            get
            {
                SharedResources.LogMessage("SupportedActions Get", "Returning empty arraylist");
                return new ArrayList();
            }
        }

        public string Action(string actionName, string actionParameters)
        {
            LogMessage("", "Action {0}, parameters {1} not implemented", actionName, actionParameters);
            throw new ASCOM.ActionNotImplementedException("Action " + actionName + " is not implemented by this driver");
        }

        public void CommandBlind(string command, bool raw)
        {
            //CheckConnected("CommandBlind");
            SharedResources.LogMessage("CommandBlind", "Command: {0}, raw: {1}", command, raw.ToString());

            if (raw)
            {
                lock (lockObject)
                {
                    SharedResources.SendSerialMessageBlind(command);
                }
            }
            else
            {
                switch (command)
                {
                    default:
                        throw new MethodNotImplementedException("MethodNotImplementedException: CommandBlind");
                }
            }
        }

        public bool CommandBool(string command, bool raw)
        {
            //CheckConnected("CommandBool");
            //SharedResources.LogMessage(MotorID.ToString() + "::CommandBool", "Command: {0}, raw: {1}", command, raw.ToString());

            if (raw)
            {
                return Convert.ToBoolean(CommandString(command, true));
            }
            else
            {
                switch (command)
                {
                    default:
                        throw new MethodNotImplementedException("MethodNotImplementedException: CommandBool");
                }
            }
        }


        public string CommandString(string command, bool raw)
        {
            //CheckConnected("CommandString");
            //SharedResources.LogMessage(MotorID.ToString() + "::CommandString", "Command: {0}, raw: {1}", command, raw.ToString());

            if (raw)
            {
                lock (lockObject)
                {
                    return SharedResources.SendSerialMessage(command);
                }
            }
            else
            {
                switch (command)
                {
                    default:
                        throw new MethodNotImplementedException("MethodNotImplementedException: CommandString");
                }
            }
        }

        public void Dispose()
        {
            // Clean up the tracelogger and util objects
            utilities.Dispose();
            utilities = null;
            astroUtilities.Dispose();
            astroUtilities = null;
        }

        public bool Connected
        {
            get
            {
                return connectedState;
            }
            set
            {
                if (value == connectedState) { return; }

                if (value)
                {
                    SharedResources.LogMessage("Connected", "Starting a new serial connection");

                    // Check if we are the first client using the shared serial
                    if (SharedResources.Connections == 0)
                    {
                        SharedResources.LogMessage("Connected", "We are the first connected client, setting serial port name");
                    }

                    SharedResources.Connected = true;

                    try
                    {
                        //
                        // Loops until the firmware reports a successful connection 
                        //
                        bool ready = false;
                        string hex = "";

                        do
                        {
                            try
                            {
                                lock (lockObject)
                                {
                                    hex = SharedResources.SendSerialMessage(GET_VERSION);
                                }
                                if (hex.Length > 0) { ready = true; }
                            }
                            catch (Exception) { ready = false; }

                        } while (!ready);

                        SharedResources.LogMessage("Connected", "Firmware ready, version " + SharedResources.ExtractNumber(hex));
                    }
                    catch (Exception e)
                    {
                        SharedResources.LogMessage("Connected", "Exception: {0}", e.Message + "\n" + e.StackTrace);
                        connectedState = false;
                        return;
                    }

                    // Check if we are the first client using the shared serial
                    if (SharedResources.Connections == 1)
                    {
                        InitDevice();
                    }

                    connectedState = true;
                    SharedResources.LogMessage("Connected", "Connected successfully");
                }
                else
                {
                    SharedResources.LogMessage("Connected", "Disconnecting the serial connection");

                    connectedState = false;
                    SharedResources.Connected = false;
                    SharedResources.LogMessage("Connected", "Disconnected successfully");
                }
            }
        }

        private void InitDevice()
        {
            try
            {
                string hex;
                string focuserVersion = "";
                lock (lockObject)
                {
                    focuserVersion = SharedResources.ExtractNumber(SharedResources.SendSerialMessage(START_ADC_READING)).ToString();
                }
                SharedResources.LogMessage("InitDevice", "Focuser version is " + focuserVersion);

                lock (lockObject)
                {
                    hex = SharedResources.SendSerialMessage(GET_CURRENT_POS);
                }
                focuserPosition = Int32.Parse(SharedResources.ExtractNumber(hex));
                SharedResources.LogMessage("InitDevice", "Current position is " + focuserPosition);

                string powerModule = "";
                lock (lockObject)
                {
                    powerModule = SharedResources.ExtractNumber(SharedResources.SendSerialMessage(POWER_MODULE));
                }
                SharedResources.LogMessage("InitDevice", "Power module is " + focuserPosition);

                string backlash = "";
                lock (lockObject)
                {
                    backlash = SharedResources.ExtractNumber(SharedResources.SendSerialMessage(BACKLASH));
                }
                char backlashMode = backlash[0];
                string backlashModeEnglish = "";
                if (backlashMode == '1') backlashModeEnglish = "None";
                else
                if (backlashMode == '2') backlashModeEnglish = "IN";
                else
                if (backlashMode == '3') backlashModeEnglish = "OUT";
                string backlashAmount = backlash.Substring(1);
                SharedResources.LogMessage("InitDevice", "Backlash is " + backlashModeEnglish + " by " + backlashAmount);

                int maxtravel = 0;
                lock(lockObject)
                {
                    maxtravel = SharedResources.ParseNumberAsInt(SharedResources.SendSerialMessage(MAXIMUM_TRAVEL));
                }
                SharedResources.LogMessage("InitDevice", "Max travel is " + maxtravel);

                string configuration = "";
                lock (lockObject)
                {
                    configuration = SharedResources.ExtractNumber(SharedResources.SendSerialMessage(CONFIGURATION));
                }
                SharedResources.LogMessage("InitDevice", "Configuration is " + configuration);
                byte dutyCycle = (byte)configuration[3];
                byte stepDelay = (byte)configuration[4];
                byte stepSize = (byte)configuration[5];
                SharedResources.LogMessage("InitDevice", "Duty Cycle is " + dutyCycle);
                SharedResources.LogMessage("InitDevice", "Step Delay is " + stepDelay);
                SharedResources.LogMessage("InitDevice", "Step Size is " + stepSize);

                lock (lockObject)
                {
                    hex = SharedResources.SendSerialMessage(GET_TEMPERATURE);
                }
                LastTemperature = SharedResources.ConvertTemperature(SharedResources.ParseNumberAsDouble(hex));
                SharedResources.LogMessage("InitDevice", "Current temperature is " + LastTemperature);
            }
            catch(Exception e)
            {
                SharedResources.LogMessage("InitDevice", "Error: " + e.Message + "\n" + e.StackTrace);
                throw;
            }
        }

        public string Description
        {
            // TODO customise this device description
            get
            {
                SharedResources.LogMessage("Description Get", driverDescription);
                return driverDescription;
            }
        }

        public string DriverInfo
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                // TODO customise this driver description
                string driverInfo = "Robofocus ASCOM driver. Version: " + String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                SharedResources.LogMessage("DriverInfo Get", driverInfo);
                return driverInfo;
            }
        }

        public string DriverVersion
        {
            get
            {
                Version version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                string driverVersion = String.Format(CultureInfo.InvariantCulture, "{0}.{1}", version.Major, version.Minor);
                SharedResources.LogMessage("DriverVersion Get", driverVersion);
                return driverVersion;
            }
        }

        public short InterfaceVersion
        {
            // set by the driver wizard
            get
            {
                LogMessage("InterfaceVersion Get", "2");
                return Convert.ToInt16("2");
            }
        }

        public string Name
        {
            get
            {
                string name = "NGCAT Focuser";
                SharedResources.LogMessage("Name Get", name);
                return name;
            }
        }

        #endregion

        #region IFocuser Implementation

        internal static int focuserPosition = 0; // Class level variable to hold the current focuser position
        private const int focuserSteps = 10000;
        private double LastTemperature { get; set; }

        public bool Absolute
        {
            get
            {
                SharedResources.LogMessage("Absolute Get", true.ToString());
                return true; // This is an absolute focuser
            }
        }

        public void Halt()
        {
            SharedResources.LogMessage("Halt", "Not implemented");
            CheckConnected("Halt");
            SharedResources.SendSerialMessageBlind(STOP_MOTION);
            SharedResources.LogMessage("Halt", "Stop motor movement");
        }

        public bool IsMoving
        {
            get
            {
                SharedResources.LogMessage("IsMoving", "Checking is focuser is moving");
                CheckConnected("IsMoving");

                return isMoving;
                /*
                bool moving = false;
                int value = 0;
                string hex;

                try
                {
                    hex = SharedResources.Receive();
                    SharedResources.LogMessage("IsMoving", "Read: " + hex);
                    if (hex.Length > 0)
                    {
                        string command = SharedResources.ParseCommand(hex);
                        if (command == "I" || command == "O")
                        {
                            SharedResources.LogMessage("IsMoving", "Moving...");
                            moving = true;
                        }
                        else
                        if (command == "FD")
                        {
                            value = Convert.ToInt16(SharedResources.ParseNumberAsInt(hex));
                            SharedResources.LogMessage("IsMoving", "Done moving. Current position is " + value);
                            moving = false;
                        }
                        else
                        {
                            SharedResources.LogMessage("IsMoving", "Not moving");
                            moving = false;
                        }
                    }
                }
                catch (Exception e)
                {
                    SharedResources.LogMessage("IsMoving", "Error: " + e.Message + "\n" + e.StackTrace);
                }

                return moving;
                */
            }
        }

        public bool Link
        {
            get
            {
                SharedResources.LogMessage("Link Get", this.Connected.ToString());
                return this.Connected; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
            set
            {
                SharedResources.LogMessage("Link Set", value.ToString());
                SharedResources.COMPortName = comPort;
                this.Connected = value; // Direct function to the connected method, the Link method is just here for backwards compatibility
            }
        }

        public int MaxIncrement
        {
            get
            {
                SharedResources.LogMessage("MaxIncrement Get", focuserSteps.ToString());
                return focuserSteps; // Maximum change in one move
            }
        }

        public int MaxStep
        {
            get
            {
                SharedResources.LogMessage("MaxStep Get", focuserSteps.ToString());
                return focuserSteps; // Maximum extent of the focuser, so position range is 0 to 10,000
            }
        }

        public void Move(int Position)
        {
            SharedResources.LogMessage("Move", "Move motor to '{0}'", Position);

            CheckConnected("Move");

            if (Position > MaxStep) { Position = MaxStep; }
            if (Position == focuserPosition) return;

            string command = SET_TARGET_POS + Position;

            isMoving = true;
            lock (lockObject)
            {
                SharedResources.SendSerialMessageBlind(command);
                string read = "";
                while (read != "F")
                {
                    Thread.Sleep(100);
                    read = SharedResources.ReceiveN(1);
                    SharedResources.LogMessage("Move", "Waiting for focuser to stop moving. Read: " + read);
                    if (read[0] == 'I') focuserPosition--;
                    else
                    if (read[0] == 'O') focuserPosition++;
                }
                if (read == "F")
                {
                    read += SharedResources.ReceiveN(8);
                    SharedResources.LogMessage("Move", "New positioin is: " + SharedResources.ParseNumberAsInt(read));
                }
            }
            isMoving = false;

            SharedResources.LogMessage("Move", "Move done");
        }


        public int Position
        {
            get
            {
                SharedResources.LogMessage("Position", "Getting current position");
                CheckConnected("Position");

                string hex = "";
                int value = 0;

                try
                {
                    lock (lockObject)
                    {
                        hex = SharedResources.SendSerialMessage(GET_CURRENT_POS);
                    }
                    value = SharedResources.ParseNumberAsInt(hex);
                }
                catch (Exception)
                {
                    //throw new FormatException("Invalid hex string received");
                    SharedResources.LogMessage("Position", "Invalid hex string received " + hex);
                    return focuserPosition;
                }

                //LogMessage("Position", "Return value int:{0} hex:{1}", value, hex);
                focuserPosition = value;
                return focuserPosition;
            }
        }


        public double StepSize
        {
            get
            {
                SharedResources.LogMessage("StepSize Get", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("StepSize", false);
            }
        }

        public bool TempComp
        {
            get
            {
                SharedResources.LogMessage("TempComp Get", false.ToString());
                return false;
            }
            set
            {
                SharedResources.LogMessage("TempComp Set", "Not implemented");
                throw new ASCOM.PropertyNotImplementedException("TempComp", false);
            }
        }

        public bool TempCompAvailable
        {
            get
            {
                SharedResources.LogMessage("TempCompAvailable Get", false.ToString());
                return false; // Temperature compensation is not available in this driver
            }
        }

        public double Temperature
        {
            get
            {
                CheckConnected("Temperature");

                double f;
                double temp = 0;
                string hex = "";
                try
                {
                    lock (lockObject)
                    {
                        hex = SharedResources.SendSerialMessage(GET_TEMPERATURE);
                    }
                    temp = SharedResources.ParseNumberAsDouble(hex);
                }
                catch (Exception e)
                {
                    //throw new FormatException("Invalid hex string received");
                    SharedResources.LogMessage("Temperature", "Invalid hex string received " + hex + ". " + e.Message + "\n" + e.StackTrace);
                    return LastTemperature;
                }

                SharedResources.LogMessage("Temperature", "Temperature is " + temp + " steps");
                f = SharedResources.ConvertTemperature(temp);
                LastTemperature = Math.Round(f, 1, MidpointRounding.AwayFromZero);
                SharedResources.LogMessage("Temperature", "Temperature is " + LastTemperature + "C");
                return LastTemperature;
            }
        }

        #endregion

        #region Private properties and methods
        // here are some useful properties and methods that can be used as required
        // to help with driver development

        #region ASCOM Registration

        // Register or unregister driver for ASCOM. This is harmless if already
        // registered or unregistered. 
        //
        /// <summary>
        /// Register or unregister the driver with the ASCOM Platform.
        /// This is harmless if the driver is already registered/unregistered.
        /// </summary>
        /// <param name="bRegister">If <c>true</c>, registers the driver, otherwise unregisters it.</param>
        private static void RegUnregASCOM(bool bRegister)
        {
            using (var P = new ASCOM.Utilities.Profile())
            {
                P.DeviceType = "Focuser";
                if (bRegister)
                {
                    P.Register(driverID, driverDescription);
                }
                else
                {
                    P.Unregister(driverID);
                }
            }
        }

        /// <summary>
        /// This function registers the driver with the ASCOM Chooser and
        /// is called automatically whenever this class is registered for COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is successfully built.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During setup, when the installer registers the assembly for COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually register a driver with ASCOM.
        /// </remarks>
        [ComRegisterFunction]
        public static void RegisterASCOM(Type t)
        {
            RegUnregASCOM(true);
        }

        /// <summary>
        /// This function unregisters the driver from the ASCOM Chooser and
        /// is called automatically whenever this class is unregistered from COM Interop.
        /// </summary>
        /// <param name="t">Type of the class being registered, not used.</param>
        /// <remarks>
        /// This method typically runs in two distinct situations:
        /// <list type="numbered">
        /// <item>
        /// In Visual Studio, when the project is cleaned or prior to rebuilding.
        /// For this to work correctly, the option <c>Register for COM Interop</c>
        /// must be enabled in the project settings.
        /// </item>
        /// <item>During uninstall, when the installer unregisters the assembly from COM Interop.</item>
        /// </list>
        /// This technique should mean that it is never necessary to manually unregister a driver from ASCOM.
        /// </remarks>
        [ComUnregisterFunction]
        public static void UnregisterASCOM(Type t)
        {
            RegUnregASCOM(false);
        }

        #endregion

        /// <summary>
        /// Returns true if there is a valid connection to the driver hardware
        /// </summary>
        private bool IsConnected
        {
            get
            {
                // TODO check that the driver hardware connection exists and is connected to the hardware
                return connectedState;
            }
        }

        /// <summary>
        /// Use this function to throw an exception if we aren't connected to the hardware
        /// </summary>
        /// <param name="message"></param>
        private void CheckConnected(string message)
        {
            if (!IsConnected)
            {
                throw new ASCOM.NotConnectedException(message);
            }
        }

        /// <summary>
        /// Read the device configuration from the ASCOM Profile store
        /// </summary>
        internal void ReadProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Focuser";
                SharedResources.TraceEnabled = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, string.Empty, traceStateDefault));
                comPort = driverProfile.GetValue(driverID, comPortProfileName, string.Empty, comPortDefault);
            }
        }

        /// <summary>
        /// Write the device configuration to the  ASCOM  Profile store
        /// </summary>
        internal void WriteProfile()
        {
            using (Profile driverProfile = new Profile())
            {
                driverProfile.DeviceType = "Focuser";
                driverProfile.WriteValue(driverID, traceStateProfileName, SharedResources.TraceEnabled.ToString());
                driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString());
            }
        }

        internal static int Sync(int position)
        {
            int temp = -1;
            try
            {
                lock (lockObject)
                {
                    temp = SharedResources.ParseNumberAsInt(SharedResources.SendSerialMessage(SYNC + position));
                }
                SharedResources.LogMessage("Sync", "Current position set to " + temp);
            }
            catch (Exception e)
            {
                SharedResources.LogMessage("Sync", "Error: " + e.Message + "\n" + e.StackTrace);
            }
            return temp;
        }

            /// <summary>
            /// Log helper function that takes formatted strings and arguments
            /// </summary>
            /// <param name="identifier"></param>
            /// <param name="message"></param>
            /// <param name="args"></param>
        internal static void LogMessage(string identifier, string message, params object[] args)
        {
            var msg = string.Format(message, args);
            SharedResources.LogMessage(identifier, msg);
        }
        #endregion
    }
}
