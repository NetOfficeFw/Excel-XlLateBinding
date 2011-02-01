using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;
 
namespace Example3
{
    public partial class Form1 : Form
    {
        XlApplication excelApplication;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
             

            // start excel and turn off msg boxes
            excelApplication = new XlApplication();
            excelApplication.DisplayAlerts = false;

            // add a new workbook
            XlWorkbook workBook = excelApplication.Workbooks.Add();
            XlWorksheet workSheet = workBook.Worksheets[1];

            // some kind of numerics

            string Pattern1 = string.Format("0{0}00", XlLateBindingApiSettings.XlThreadCulture.NumberFormat.CurrencyDecimalSeparator);
            string Pattern2 = string.Format("#{1}##0{0}00", XlLateBindingApiSettings.XlThreadCulture.NumberFormat.CurrencyDecimalSeparator, XlLateBindingApiSettings.XlThreadCulture.NumberFormat.CurrencyGroupSeparator);
                
            workSheet.Range("A1").Value = "Type";
            workSheet.Range("B1").Value = "Value";
            workSheet.Range("C1").Value = "Formatted " + Pattern1;
            workSheet.Range("D1").Value = "Formatted " + Pattern2;

            int integerValue = 532234;            
            workSheet.Range("A3").Value = "Integer";
            workSheet.Range("B3").Value = integerValue;
            workSheet.Range("C3").Value = integerValue;
            workSheet.Range("C3").NumberFormat = Pattern1;
            workSheet.Range("D3").Value = integerValue;
            workSheet.Range("D3").NumberFormat = Pattern2;

            double doubleValue = 23172.64;
            workSheet.Range("A4").Value = "double";
            workSheet.Range("B4").Value = doubleValue;
            workSheet.Range("C4").Value = doubleValue;
            workSheet.Range("C4").NumberFormat = Pattern1;
            workSheet.Range("D4").Value = doubleValue;
            workSheet.Range("D4").NumberFormat = Pattern2;

            float floatValue = 84345.9132f;
            workSheet.Range("A5").Value = "float";
            workSheet.Range("B5").Value = floatValue;
            workSheet.Range("C5").Value = floatValue;
            workSheet.Range("C5").NumberFormat = Pattern1;
            workSheet.Range("D5").Value = floatValue;
            workSheet.Range("D5").NumberFormat = Pattern2;

            Decimal decimalValue = 7251231.313367m;
            workSheet.Range("A6").Value = "Decimal";
            workSheet.Range("B6").Value = decimalValue;
            workSheet.Range("C6").Value = decimalValue;
            workSheet.Range("C6").NumberFormat = Pattern1;
            workSheet.Range("D6").Value = decimalValue;
            workSheet.Range("D6").NumberFormat = Pattern2;

            workSheet.Range("A9").Value = "DateTime";
            workSheet.Range("B10").Value = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.FullDateTimePattern;
            workSheet.Range("C10").Value = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.LongDatePattern;
            workSheet.Range("D10").Value = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.ShortDatePattern;
            workSheet.Range("E10").Value = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.LongTimePattern;
            workSheet.Range("F10").Value = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.ShortTimePattern;

            // DateTime

            DateTime dateTimeValue = DateTime.Now;
            workSheet.Range("B11").Value = dateTimeValue;
            workSheet.Range("B11").NumberFormat = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.FullDateTimePattern;

            workSheet.Range("C11").Value = dateTimeValue;
            workSheet.Range("C11").NumberFormat = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.LongDatePattern;

            workSheet.Range("D11").Value = dateTimeValue;
            workSheet.Range("D11").NumberFormat = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.ShortDatePattern;

            workSheet.Range("E11").Value = dateTimeValue;
            workSheet.Range("E11").NumberFormat = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.LongTimePattern;
            
            workSheet.Range("F11").Value = dateTimeValue;
            workSheet.Range("F11").NumberFormat = XlLateBindingApiSettings.XlThreadCulture.DateTimeFormat.ShortTimePattern;

            // string
            workSheet.Range("A14").Value = "String";
            workSheet.Range("B14").Value = "This is a sample String";
            workSheet.Range("B14").NumberFormat = "@";
            // number as string
            workSheet.Range("B15").Value = "513";
            workSheet.Range("B15").NumberFormat = "@";

            // set colums
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
       
            // save the book 
            string fileExtension = XlConverter.GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example3{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }
    }
}
