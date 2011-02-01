using System;
using System.Drawing;
using System.Windows.Forms;

using Excel = LateBindingApi.Excel;
using LateBindingApi.Excel.Enums;

namespace Example6
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
            
            // dont show dialogs with an invisible excel
            _excelApplication.Visible = true;

            // add a new workbook
            Excel.Workbook workBook = _excelApplication.Workbooks.Add();
            Excel.Worksheet workSheet = workBook.Worksheets[1];

            // show selected window and display user clicks ok or cancel
            bool returnValue = false;
            RadioButton radioSelectButton = GetSelectedRadioButton();
            switch (radioSelectButton.Text)
            {
                case "xlDialogAddinManager":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogAddinManager].Show();
                    break;

                case "xlDialogFont":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogFont].Show();
                    break;

                case "xlDialogEditColor":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogEditColor].Show();
                    break;

                case "xlDialogGallery3dBar":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogGallery3dBar].Show();
                    break;

                case "xlDialogSearch":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogSearch].Show();
                    break;

                case "xlDialogPrinterSetup":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogPrinterSetup].Show();
                    break;

                case "xlDialogFormatNumber":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogFormatNumber].Show();
                    break;

                case "xlDialogApplyStyle":

                    returnValue = _excelApplication.Dialogs[XlBuiltInDialog.xlDialogApplyStyle].Show();
                    break;

                default:
                    throw (new Exception("Unkown dialog selected."));

            }

            string message = string.Format("The dialog returns {0}.", returnValue);
            MessageBox.Show(this, message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
  
            // close excel and dispose reference
            _excelApplication.Quit();
            _excelApplication.Dispose();

        }
        
        #region Helper
        
        private RadioButton GetSelectedRadioButton()
        {
            foreach (Control itemControl in panelSelection.Controls)
            {
                RadioButton radioSelectButton = itemControl as RadioButton;
                if (null != radioSelectButton)
                {
                    if (radioSelectButton.Checked)
                        return radioSelectButton;
                }
            }

            throw (new Exception("No Dialog selected."));
        }

        #endregion
    }
}
