using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;
using LateBindingApi.Excel.VBIDE;
    
namespace Example7
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
            
            // add new global Code Module
            XlVBComponent globalModule = workBook.VBProject.VBComponents.Add(LateBindingApi.Excel.Enums.vbext_ComponentType.vbext_ct_StdModule);
            globalModule.Name = "MyNewCodeModule";

            // add a new procedure to the modul
            globalModule.CodeModule.InsertLines(1, "Public Sub HelloWorld(Param as string)\r\n MsgBox \"Hello World!\" & vbnewline & Param\r\nEnd Sub");

            // create a click event trigger for the first worksheet
            int linePosition = workBook.VBProject.VBComponents[2].CodeModule.CreateEventProc("BeforeDoubleClick", "Worksheet");
            workBook.VBProject.VBComponents[2].CodeModule.InsertLines(linePosition + 1, "HelloWorld \"BeforeDoubleClick\"");

            // display info in the worksheet
            workBook.Worksheets[1].Cells(2, 2).Value = "This workbook contains dynamic created VBA Moduls and Event Code";
            workBook.Worksheets[1].Cells(5, 2).Value = "Open the VBA Editor to see the code";
            workBook.Worksheets[1].Cells(8, 2).Value = "Do a double click to catch the BeforeDoubleClick Event from this Worksheet.";

            // save the book 
            string fileExtension = XlConverter.GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example7{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile);

             // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }
    }
}
