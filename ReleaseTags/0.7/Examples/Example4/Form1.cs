using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Shapes;

namespace Example4
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

            workSheet.Cells(1, 1).Value = "these sample shapes was dynamicly created by code.";

            // create a star
            XlShape starShape = workSheet.Shapes.AddShape(LateBindingApi.Excel.Enums.MsoAutoShapeType.msoShape32pointStar, 10, 50, 200, 20);
            
            // create a simple textbox
            XlShape textBox = workSheet.Shapes.AddTextbox(LateBindingApi.Excel.Enums.MsoTextOrientation.msoTextOrientationHorizontal, 10, 150, 200, 50);
            textBox.TextFrame.Characters().Text = "text";
            textBox.TextFrame.Characters().Font.Size = 14;

            // create a wordart
            XlShape textEffect =  workSheet.Shapes.AddTextEffect(LateBindingApi.Excel.Enums.MsoPresetTextEffect.msoTextEffect14,  "WordArt", "Arial", 12, LateBindingApi.Excel.Enums.MsoTriState.msoTrue, LateBindingApi.Excel.Enums.MsoTriState.msoFalse, 10, 250);

            // create text effect
            XlShape textDiagram = workSheet.Shapes.AddTextEffect(LateBindingApi.Excel.Enums.MsoPresetTextEffect.msoTextEffect11, "Effect", "Arial", 14, LateBindingApi.Excel.Enums.MsoTriState.msoFalse, LateBindingApi.Excel.Enums.MsoTriState.msoFalse ,10, 350);    

            // save the book 
            string fileExtension = XlConverter.GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example4{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile);

            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }
    }
}
