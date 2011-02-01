using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

using Excel = TcKs.MSOffice.Excel;

namespace OfficeForNet
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("OfficeFor.Net Performance Test - 5000 Cells.");

            /*
             * start excel and disable messageboxes and screen updating
             */
            Excel.Application excelApplication = Excel.Application.CreateApplication();
            // no propertiy support for excel object in current release 0_1_2
            // we do it manualy
            object[] paramArray = new object[1];
            paramArray[0] = false;
            excelApplication.GetWrappedObject().GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, excelApplication.GetWrappedObject(), paramArray, null);
            excelApplication.GetWrappedObject().GetType().InvokeMember("ScreenUpdating", BindingFlags.SetProperty, null, excelApplication.GetWrappedObject(), paramArray, null);

            /*
            *  create new empty worksheet
            */
            Excel.Workbooks books = excelApplication.Workbooks;
            Excel.Workbook book = books.Add();
            Excel.Worksheets sheets = book.Worksheets;
            Excel.Worksheet sheet = (Excel.Worksheet)sheets.Add();

            /*
            *  do the test
            */
            DateTime timeStart = DateTime.Now;
            for (int i = 1; i <= 5000; i++)
            {
                // cells property for a sheet are also not supported :-(
                string rangeAdress = "$A" +i.ToString() ;
                Excel.Range cellRange = (Excel.Range)sheet.GetRange(rangeAdress);
                cellRange.Value2 = "value";
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
            string logFile = Path.Combine(Environment.CurrentDirectory, "OfficeForNet.log");
            string logFileAppend = timeElapsed.ToString() + Environment.NewLine;
            File.AppendAllText(logFile, logFileAppend, Encoding.UTF8);
            
            excelApplication.Quit();          
            TcKs.MSOffice.Common.WrapperHelper.GCCollect();
        }
    }
}
