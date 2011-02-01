using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;

namespace Example2
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
           
            // font action
            workSheet.Range("A1").Value = "Arial Size:8 Bold Italic Underline";
            workSheet.Range("A1").Font.Name = "Arial";
            workSheet.Range("A1").Font.Size = 8;
            workSheet.Range("A1").Font.Bold = true;
            workSheet.Range("A1").Font.Italic = true;
            workSheet.Range("A1").Font.Underline = true;
            workSheet.Range("A1").Font.Color = Color.Violet.ToArgb();

            workSheet.Range("A3").Value = "Times New Roman Size:10";
            workSheet.Range("A3").Font.Name = "Times New Roman";
            workSheet.Range("A3").Font.Size = 10;
            workSheet.Range("A3").Font.Color = Color.Orange.ToArgb();

            workSheet.Range("A5").Value = "Comic Sans MS Size:12 WrapText";
            workSheet.Range("A5").Font.Name = "Comic Sans MS";
            workSheet.Range("A5").Font.Size = 12;
            workSheet.Range("A5").WrapText = true;
            workSheet.Range("A5").Font.Color = Color.Navy.ToArgb();
            
            // HorizontalAlignment
            workSheet.Range("A7").Value = "xlHAlignLeft";
            workSheet.Range("A7").HorizontalAlignment = LateBindingApi.Excel.Enums.XlHAlign.xlHAlignLeft;

            workSheet.Range("B7").Value = "xlHAlignCenter";
            workSheet.Range("B7").HorizontalAlignment = LateBindingApi.Excel.Enums.XlHAlign.xlHAlignCenter;

            workSheet.Range("C7").Value = "xlHAlignRight";
            workSheet.Range("C7").HorizontalAlignment = LateBindingApi.Excel.Enums.XlHAlign.xlHAlignRight;

            workSheet.Range("D7").Value = "xlHAlignJustify";
            workSheet.Range("D7").HorizontalAlignment = LateBindingApi.Excel.Enums.XlHAlign.xlHAlignJustify;

            workSheet.Range("E7").Value = "xlHAlignDistributed";
            workSheet.Range("E7").HorizontalAlignment = LateBindingApi.Excel.Enums.XlHAlign.xlHAlignDistributed;

            workSheet.Range("A9").Value = "xlVAlignTop";
            workSheet.Range("A9").VerticalAlignment = LateBindingApi.Excel.Enums.XlVAlign.xlVAlignTop;

            workSheet.Range("B9").Value = "xlVAlignCenter";
            workSheet.Range("B9").VerticalAlignment = LateBindingApi.Excel.Enums.XlVAlign.xlVAlignCenter;

            workSheet.Range("C9").Value = "xlVAlignBottom";
            workSheet.Range("C9").VerticalAlignment = LateBindingApi.Excel.Enums.XlVAlign.xlVAlignBottom;

            workSheet.Range("D9").Value = "xlVAlignDistributed";
            workSheet.Range("D9").VerticalAlignment = LateBindingApi.Excel.Enums.XlVAlign.xlVAlignDistributed;

            workSheet.Range("E9").Value = "xlVAlignJustify";
            workSheet.Range("E9").VerticalAlignment = LateBindingApi.Excel.Enums.XlVAlign.xlVAlignJustify;

            // setup rows and columns
            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].ColumnWidth = 25;
            workSheet.Columns[3].ColumnWidth = 25;
            workSheet.Columns[4].ColumnWidth = 25;
            workSheet.Columns[5].ColumnWidth = 25;
            workSheet.Rows[9].RowHeight = 35;

            // save the book 
            string fileExtension = XlConverter.GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example2{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }
    }
}
