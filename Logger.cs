using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XMLToCSV
{
    public static class Logger
    {
        public static void WriteLog(string message, string path, string logFilename)
        {

            try
            {
                //string logPath = ConfigurationManager.AppSettings["logPath"];
                string pfl = path + logFilename;

                using (StreamWriter writer = new StreamWriter(pfl, true))
                {
                    writer.WriteLine($"{DateTime.Now} ~ {message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
