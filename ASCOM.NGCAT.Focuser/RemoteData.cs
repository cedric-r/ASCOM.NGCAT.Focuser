using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using ASCOM.NGCAT;
using Newtonsoft.Json;

namespace ASCOM.NGCAT
{
    public class DataItem
    {
        public int time;
        public double temperature;
        public double humidity;
        public double dew;
        public double pressure;
        public double cloud;
        public int rain;
        public double board;
        public double wind;
        public double gust;
        public double light;
    }

    public static class RemoteData
    {
        private static int UPDATEFREQUENCY = 120; // In seconds
        private static List<DataItem> _Data = new List<DataItem>();
        private static DateTime _LastUpdate = DateTime.MinValue;

        private static Mutex semaphore = new Mutex();

        public static List<DataItem> GetData(string server)
        {
            try
            {
                semaphore.WaitOne();
                if ((DateTime.Now - _LastUpdate).TotalSeconds > UPDATEFREQUENCY || _Data.Count == 0)
                {
                    HttpClient _Client = new HttpClient();
                    String request = server;
                    SharedResources.LogMessage("URL is " + request);

                    using (HttpResponseMessage response = _Client.GetAsync(request).Result)
                    {
                        using (HttpContent content = response.Content)
                        {
                            // ... Read the string.
                            string result = content.ReadAsStringAsync().Result;
                            SharedResources.LogMessage("Set Server result " + result);
                            _Data = new List<DataItem>();
                            _Data.Add(Decoder(result));
                            _LastUpdate = DateTime.Now;
                        }
                    }
                }

                return _Data;
            }
            catch (Exception e)
            {
                SharedResources.LogMessage("Load data error " + e.ToString());
                throw;
            }
            finally
            {
                semaphore.ReleaseMutex();
            }
        }

        private static DataItem Decoder(string data)
        {
            SharedResources.LogMessage("Data=" + data);
            DataItem di = new DataItem();
            List<string> diList = data.Split('\n').ToList();
            foreach (string line in diList)
            {
                if (line.Contains("clouds=")) di.cloud = ConvertToDouble(line.Replace("clouds=", ""));
                if (line.Contains("temp=")) di.temperature = ConvertToDouble(line.Replace("temp=", ""));
                if (line.Contains("rain="))
                {
                    double r = ConvertToDouble(line.Replace("rain=", ""));
                    if (r > 2000) di.rain = 1;
                    else di.rain = 0;
                }

                if (line.Contains("wind="))
                {
                    di.wind = ConvertToDouble(line.Replace("wind=", ""));
                    if (di.wind != -1) di.wind = (double)Math.Truncate((di.wind / 3.6) * 100) / 100;
                }

                if (line.Contains("gust="))
                {
                    di.gust = ConvertToDouble(line.Replace("gust=", ""));
                    if (di.gust != -1) di.gust = (double)Math.Truncate((di.gust / 3.6) * 100) / 100;
                }

                if (line.Contains("light=")) di.light = ConvertToDouble(line.Replace("light=", ""));
                if (line.Contains("hum=")) di.humidity = ConvertToDouble(line.Replace("hum=", ""));
                if (line.Contains("dewp=")) di.dew = ConvertToDouble(line.Replace("dewp=", ""));

            }
            SharedResources.LogMessage("DataItem=" + JsonConvert.SerializeObject(di));
            return di;
        }

        private static double ConvertToDouble(string num)
        {
            string a = Convert.ToString(Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator);
            num = num.Replace(".", a);
            num = num.Replace(",", a);
            return Double.Parse(num);
        }


    }

}