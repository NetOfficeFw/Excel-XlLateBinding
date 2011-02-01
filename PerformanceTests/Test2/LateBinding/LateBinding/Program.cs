using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Enums;

namespace LateBinding
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("XlLateBinding Performance Test - 5000 Cells.");

            /*
             * start excel and disable messageboxes and screen updating
             */
            XlApplication excelApplication = new XlApplication();
            excelApplication.DisplayAlerts = false;
            excelApplication.ScreenUpdating = false;

            /*
            *  create new empty worksheet
            */
            excelApplication.Workbooks.Add();
            XlWorksheet sheet = excelApplication.Workbooks[1].Worksheets[1];

            /*
            *  do the test
            */
            DateTime timeStart = DateTime.Now;
            for (int i = 1; i <= 5000; i++)
            {
                // cells property for a sheet in OfficeFor.Net are not supported
                // the reason for all examples use range
                string rangeAdress = "$A" + i.ToString();
                XlRange cellRange = sheet.Range(rangeAdress);
                cellRange.Value = "value";
                cellRange.Font.Name = "Verdana";
                cellRange.NumberFormat = "@";

                cellRange.BorderAround(XlLineStyle.xlDouble, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, System.Drawing.Color.Black);
            }
            DateTime timeEnd = DateTime.Now;
            TimeSpan timeElapsed = timeEnd - timeStart;

            /*
            * display for user
            */
            string outputConsole = string.Format("Time Elapsed: {0}{1}Press any Key.", timeElapsed, Environment.NewLine);
            Console.WriteLine(outputConsole);
            Console.Read();

            /*
           * write result in logfile
           */
            string logFile = Path.Combine(Environment.CurrentDirectory, "LateBinding.log");
            string logFileAppend = timeElapsed.ToString() + Environment.NewLine;
            File.AppendAllText(logFile, logFileAppend, Encoding.UTF8);

            excelApplication.Quit();
            excelApplication.Dispose();
        }
    }
}
