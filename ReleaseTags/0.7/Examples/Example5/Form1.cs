using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Charts;

namespace Example5
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

            // we need some data to display
            XlRange dataRange = PutSampleData(workSheet);
            
            // create a nice diagram
            XlChartObject chart = workSheet.ChartObjects.Add(70, 100, 375, 225);
            chart.Chart.SetSourceData(dataRange);

            // save the book 
            string fileExtension = XlConverter.GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example5{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile);
             
            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }

        private XlRange PutSampleData(XlWorksheet workSheet)
        {   
            workSheet.Cells(2, 2).Value = "Datum";
            workSheet.Cells(3, 2).Value = DateTime.Now.ToShortDateString();
            workSheet.Cells(4, 2).Value = DateTime.Now.ToShortDateString();
            workSheet.Cells(5, 2).Value = DateTime.Now.ToShortDateString();
            workSheet.Cells(6, 2).Value = DateTime.Now.ToShortDateString();


            workSheet.Cells(2, 3).Value = "Column1";
            workSheet.Cells(3, 3).Value = 25;
            workSheet.Cells(4, 3).Value = 33;
            workSheet.Cells(5, 3).Value = 30;
            workSheet.Cells(6, 3).Value = 22;

            workSheet.Cells(2, 4).Value = "Column2";
            workSheet.Cells(3, 4).Value = 25;
            workSheet.Cells(4, 4).Value = 33;
            workSheet.Cells(5, 4).Value = 30;
            workSheet.Cells(6, 4).Value = 22;

            workSheet.Cells(2, 5).Value = "Column3";
            workSheet.Cells(3, 5).Value = 25;
            workSheet.Cells(4, 5).Value = 33;
            workSheet.Cells(5, 5).Value = 30;
            workSheet.Cells(6, 5).Value = 22;

            return workSheet.Range("$B2:$E6");
        }

    }
}
