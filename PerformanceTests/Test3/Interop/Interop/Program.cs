using System;
using System.IO;
using System.Collections.Generic; 
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Excel;

namespace Interop
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Interop Performance Test - 10.000 Cells.");

            /*
             * start excel and disable messageboxes and screen updating
             */
            Excel.Application excelApplication = new Excel.Application();
            excelApplication.DisplayAlerts = false;
            excelApplication.ScreenUpdating  = false;
            //excelApplication.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(excelApplication_WorkbookActivate);
            /*
            *  create new empty worksheet
            */
            Excel.Workbooks books = excelApplication.Workbooks;
            Excel.Workbook  book = books.Add(Missing.Value);
            Excel.Sheets    sheets = book.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)sheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            /*
            *  do the test
            *  we collect all references and release after time recording
            *  the 2 latebind libs release the references at end automaticly and we want a fair test
            */
            List<object> comReferenesList = new List<object>();
            DateTime timeStart = DateTime.Now;
            for (int i = 1; i <= 10000; i++)
            {
                string rangeAdress = "$A" + i.ToString();
                Range cellRange = (Range)sheet.Range[rangeAdress];
                cellRange.Value = "value";
                cellRange.Font.Name = "Verdana";
                cellRange.NumberFormat = "@";
                cellRange.WrapText = false;                
                Comment sampleComment = cellRange.AddComment("Sample Comment");        
                comReferenesList.Add(cellRange);
                comReferenesList.Add(sampleComment);
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
            string logFile = Path.Combine(Environment.CurrentDirectory, "Interop.log");
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
