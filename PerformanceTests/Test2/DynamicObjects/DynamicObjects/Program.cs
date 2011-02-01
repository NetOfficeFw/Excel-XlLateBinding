using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.IO;

namespace DynamicObjects
{
    class Program
    {
  
        static void Main(string[] args)
        {
            Console.WriteLine("DynamicObjects Performance Test - 5000 Cells.");
            
            /*
             * start excel and disable messageboxes and screen updating
             */
            Type excelType = System.Type.GetTypeFromProgID("Excel.Application");
            dynamic excelApplication = System.Activator.CreateInstance(excelType);
            excelApplication.DisplayAlerts = false;
            excelApplication.ScreenUpdating = false;
           
            /*
            *  create new empty worksheet
            */
            dynamic books = excelApplication.Workbooks;
            dynamic book = books.Add(Missing.Value);
            dynamic sheets = book.Worksheets;
            dynamic sheet = sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            /*
            *  do the test
            *  we collect all references and release after time recording
            *  the 2 latebind libs release the references at end automaticly and we want a fair test
            */
            List<object> comReferenesList = new List<object>();
            DateTime timeStart = DateTime.Now;
            for (int i = 1; i <= 5000; i++)
            {
                // cells property for a sheet in OfficeForNet are not supported
                // the reason for all examples use range
                string rangeAdress = "$A" + i.ToString();
                dynamic cellRange = sheet.Range[rangeAdress];
                cellRange.Value = "value";
                cellRange.Font.Name = "Verdana";
                cellRange.NumberFormat = "@";
                cellRange.BorderAround(-4119, -4138, -4105, 0);

                comReferenesList.Add(cellRange);
            }
            DateTime timeEnd = DateTime.Now;
            TimeSpan timeElapsed = timeEnd - timeStart;

            foreach (var item in comReferenesList)
            { 
                Marshal.ReleaseComObject(item);
            }

            /*
            * display for user
            */
            string outputConsole = string.Format("Time Elapsed: {0}{1}Press any Key.", timeElapsed, Environment.NewLine);
            Console.WriteLine(outputConsole);
            Console.Read();

            /*
            * write result in logfile
            */
            string logFile = Path.Combine(Environment.CurrentDirectory, "DynamicObjects.log");
            string logFileAppend = timeElapsed.ToString() + Environment.NewLine;
            File.AppendAllText(logFile, logFileAppend, Encoding.UTF8);

            /*
            * release & quit
            */
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(books);

            excelApplication.Quit();
            Marshal.ReleaseComObject(excelApplication); 

        }
    }
}
