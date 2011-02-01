using System;
using System.Drawing;
using System.Windows.Forms;

using Excel = LateBindingApi.Excel;
using LateBindingApi.Office.Enums;

namespace Example4
{
    public partial class Form1 : Form
    {
        Excel.Application _excelApplication;

        public Form1()
        {
            InitializeComponent();

            /*
            * Initialize Api COMObject & COMVariant Support
            */
            LateBindingApi.Core.Factory.Initialize();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // start excel and turn off msg boxes
            _excelApplication = new Excel.Application();
            _excelApplication.DisplayAlerts = false;

            // add a new workbook
            Excel.Workbook workBook = _excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = workBook.Worksheets[1];

            workSheet.Cells[1,1].Value = "these sample shapes was dynamicly created by code.";

            // create a star
            Excel.Shape starShape = workSheet.Shapes.AddShape(MsoAutoShapeType.msoShape32pointStar, 10, 50, 200, 20);
            
            // create a simple textbox
            Excel.Shape textBox = workSheet.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, 150, 200, 50);
            textBox.TextFrame.Characters().Text = "text";
            textBox.TextFrame.Characters().Font.Size = 14;

            // create a wordart
            Excel.Shape textEffect = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect14, "WordArt", "Arial", 12, 
                                                                                MsoTriState.msoTrue, MsoTriState.msoFalse, 10, 250);

            // create text effect
            Excel.Shape textDiagram = workSheet.Shapes.AddTextEffect(MsoPresetTextEffect.msoTextEffect11, "Effect", "Arial", 14, 
                                                                                MsoTriState.msoFalse, MsoTriState.msoFalse, 10, 350);    

            // save the book 
            string fileExtension = GetDefaultExtension(_excelApplication);
            string workbookFile = string.Format("{0}\\Example04{1}", Environment.CurrentDirectory, fileExtension);
            workBook.SaveAs(workbookFile);

            // close excel and dispose reference
            _excelApplication.Quit();
            _excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }
        
        #region Helper
        
        /// <summary>
        /// returns the valid file extension for the instance. for example ".xls" or ".xlsx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        private static string GetDefaultExtension(Excel.Application application)
        {
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".xlsx";
            else
                return ".xls";
        }

        #endregion
    }
}
