using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Aspose.Cells;
using System.IO;

using System.Configuration;
using System.Collections;
using static XMLToCSV.Program;

namespace XMLToCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertingXLSToCSV();
                       
        }

        public static void ConvertingXLSToCSV()
        {

            //string source1 = @"D:\Sanjay\Demo\Source\New folder\HPS_Manila.xls";
            //string Destination = @"D:\Sanjay\Demo\Source\New folder\CSVFile\Result1";

            string sourceFilePath = ConfigurationManager.AppSettings["source"];
            string destination = ConfigurationManager.AppSettings["destination"];
            string archive = ConfigurationManager.AppSettings["archive"];
            string logpath = ConfigurationManager.AppSettings["logPath"];
            string currentDateTm = DateTime.Now.Date.ToString("yyyyMMdd");
            string[] allFiles = Directory.GetFiles(sourceFilePath, "*", SearchOption.TopDirectoryOnly);
            //bool result;

            try
            {
                foreach (string files in allFiles)
                {
                    FileInfo fileInfo = new FileInfo(files);
                    //file_into = file_into.Replace(files, "").Replace("xlsx","/");
                    //string filePathWithoutExt = Path.ChangeExtension(files, null);
                    string short_name = fileInfo.Name;
                    Logger.WriteLog("Start Converting xlsx to csv: " + short_name, logpath, currentDateTm+"_ConvertingFile.txt");
                    string output = Path.ChangeExtension(destination + short_name, ".csv");
                    //result = SaveAsCsv(files, output);
                    Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(files);
                    workbook.Save(output, SaveFormat.Csv);
                    
                    File.Delete(files);
                   // DeleteRow();
                }
                DeleteRow();
            }
            catch (Exception ex)
            {
                Logger.WriteLog("Error converting xlsx to csv file: " + ex, logpath, currentDateTm+ "_ErrorConvertingFile.txt");
            }
        
        }

        #region SaveAsSCV Method Commented

        //public static bool SaveAsCsv(string excelFilePath, string destinationCsvFilePath)
        //{

        //    string logpath = ConfigurationManager.AppSettings["logPath"];
        //    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        //    using (var stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        //    {

        //        try
        //        {

        //            IExcelDataReader reader = null;
        //            if (excelFilePath.EndsWith(".xls"))
        //            {
        //                reader = ExcelReaderFactory.CreateBinaryReader(stream);
        //            }
        //            else if (excelFilePath.EndsWith(".xlsx"))
        //            {
        //                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
        //            }

        //            if (reader == null)
        //                return false;

        //            var ds = reader.AsDataSet(new ExcelDataSetConfiguration()
        //            {
        //                ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
        //                {
        //                    UseHeaderRow = false
        //                }
        //            });

        //            var csvContent = string.Empty;
        //            int row_no = 0;
        //            while (row_no < ds.Tables[0].Rows.Count)
        //            {
        //                var arr = new List<string>();
        //                for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
        //                {
        //                    arr.Add(ds.Tables[0].Rows[row_no][i].ToString());
        //                }
        //                row_no++;
        //                csvContent += string.Join(",", arr) + "\n";
        //            }
        //            StreamWriter csv = new StreamWriter(destinationCsvFilePath, true);
        //            csv.Write(csvContent);
        //            csv.Close();
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex.Message);
        //            Logger.WriteLog("Some issue to reading the file" + ex.Message, logpath, "_FileReading_ErrorLogs.txt");
        //            return false;
        //        }
        //        return true;
        //    }
        //    // }
        //}
        #endregion


        public static void DeleteRow()
        {
            string currentDateTm = DateTime.Now.Date.ToString("yyyyMMdd");
            string destination = ConfigurationManager.AppSettings["afterRemovedRow"];
            string logpath = ConfigurationManager.AppSettings["logPath"];
            string sourceCSV = ConfigurationManager.AppSettings["sourceCSV"]; //@"D:\Sanjay\Demo\Source\New folder\CSVFile\";
            try
            {
                string[] allfiles = Directory.GetFiles(sourceCSV, "*", SearchOption.TopDirectoryOnly);
                List<string> linesToWrite = new List<string>();
                foreach (string file in allfiles)
                {
                 
                    string[] lines = File.ReadAllLines(file).Skip(1).ToArray();
                    var len = lines.Length;
                    var lst3 = len - 3;
                    var last2 = len - 2;
                    FileInfo fileInfo = new FileInfo(file);
                    string short_name = fileInfo.Name;
                    Logger.WriteLog("Start removing row of csv file: " + short_name, logpath, currentDateTm+"_RemovingRowCSVFile.txt");
                    string output = Path.ChangeExtension(destination + short_name, ".csv");
                    foreach (string s in lines)
                    {                        
                        linesToWrite.Add(s);

                    }
                   
                    if (len <= 4)
                    {
                        linesToWrite.RemoveRange(last2, 2);
                    }else
                    linesToWrite.RemoveRange(lst3, 3);
           
                    File.WriteAllLines(output, linesToWrite);
                    linesToWrite.Clear();
                    
                    File.Delete(file);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLog("Error removing row of csv file: "+ex, logpath, currentDateTm+"_ErrorRemovingRowCSVFile.txt");
            }

        }
              


        #region File Moving
        public static void FileMove(string allFiles)
        {
            string archive = ConfigurationManager.AppSettings["archive"];
            string logpath = ConfigurationManager.AppSettings["logPath"];
            //string logPath = ConfigurationManager.AppSettings["logPath"];// + currentDateTm;
            //string destination2 = ConfigurationManager.AppSettings["destination"];// + currentDateTm + "\\";
            //string[] files = Directory.GetFiles(sourceFilePath);
            try
            {
                if (!Directory.Exists(archive))
                {
                    Directory.CreateDirectory(archive);
                }
                //foreach (string file in allFiles)
                //{
                try
                {
                    string filname = System.IO.Path.GetFileName(allFiles);
                    archive = archive + filname;
                    Console.WriteLine(allFiles + " = " + archive);
                    // Ensure that the target does not exist.
                    if (File.Exists(archive))
                    {
                        File.Delete(archive);
                        File.Move(logpath, archive);

                    }
                    else
                    { File.Move(allFiles, archive); }
                }

                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Logger.WriteLog("Some issue to moving file" + ex.Message, logpath, "_FileMove_ErrorLogs.txt");
                }
                //}
            }
            catch (Exception ex)
            {
                Logger.WriteLog("Some issue to moving file" + ex.Message, logpath, "_FileMove_ErrorLogs.txt");
            }
        }
        #endregion

    }
}
