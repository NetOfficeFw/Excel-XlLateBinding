using System;
using System.Drawing;
using System.Windows.Forms;

using Excel = LateBindingApi.Excel;
using VBIDE = LateBindingApi.VBIDE;
using LateBindingApi.VBIDE.Enums;
    
namespace Example7
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
            try
            {
                // start excel and turn off msg boxes
                _excelApplication = new Excel.Application();
                _excelApplication.DisplayAlerts = false;

                // add a new workbook
                Excel.Workbook workBook = _excelApplication.Workbooks.Add();

                // add new global Code Module
                VBIDE.VBComponent globalModule = workBook.VBProject.VBComponents.Add(Vbext_ComponentType.vbext_ct_StdModule);
                globalModule.Name = "MyNewCodeModule";

                // add a new procedure to the modul
                globalModule.CodeModule.InsertLines(1, "Public Sub HelloWorld(Param as string)\r\n MsgBox \"Hello World!\" & vbnewline & Param\r\nEnd Sub");

                // create a click event trigger for the first worksheet
                int linePosition = workBook.VBProject.VBComponents[2].CodeModule.CreateEventProc("BeforeDoubleClick", "Worksheet");
                workBook.VBProject.VBComponents[2].CodeModule.InsertLines(linePosition + 1, "HelloWorld \"BeforeDoubleClick\"");

                // display info in the worksheet
                workBook.Worksheets[1].Cells[2, 2].Value = "This workbook contains dynamic created VBA Moduls and Event Code";
                workBook.Worksheets[1].Cells[5, 2].Value = "Open the VBA Editor to see the code";
                workBook.Worksheets[1].Cells[8, 2].Value = "Do a double click to catch the BeforeDoubleClick Event from this Worksheet.";

                // save the book 
                string fileExtension = GetDefaultExtension(_excelApplication);
                string workbookFile = string.Format("{0}\\Example07{1}", Environment.CurrentDirectory, fileExtension);
                workBook.SaveAs(workbookFile);

                FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
                fDialog.ShowDialog(this);
            }
            catch (System.Reflection.TargetInvocationException throwedException)
            {
                string message = string.Format("An error is occured.{0}ExceptionTrace:{0}", Environment.NewLine);
               
                Exception exception = throwedException;
                while (null != exception)
                {
                    message += string.Format("{0}{1}", exception.Message, Environment.NewLine); 
                    exception = exception.InnerException;
                }

                MessageBox.Show(message); 
            }
            finally
            {
                // close excel and dispose reference
                _excelApplication.Quit();
                _excelApplication.Dispose();

            }
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
