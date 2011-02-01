using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Dialogs;

namespace Example6
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
            
            // dont show dialogs with an invisible excel
            excelApplication.Visible = true;

            // add a new workbook
            XlWorkbook workBook = excelApplication.Workbooks.Add();
            XlWorksheet workSheet = workBook.Worksheets[1];

            // show selected window and display user clicks ok or cancel
            bool returnValue = false;
            RadioButton radioSelectButton = GetSelectedRadioButton();
            switch (radioSelectButton.Text)
            {
                case "xlDialogAddinManager":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogAddinManager].Show();
                    break;

                case "xlDialogFont":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogFont].Show();
                    break;

                case "xlDialogEditColor":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogEditColor].Show();
                    break;

                case "xlDialogGallery3dBar":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogGallery3dBar].Show();
                    break;

                case "xlDialogSearch":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogSearch].Show();
                    break;

                case "xlDialogPrinterSetup":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogPrinterSetup].Show();
                    break;

                case "xlDialogFormatNumber":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogFormatNumber].Show();
                    break;

                case "xlDialogApplyStyle":

                    returnValue = excelApplication.Dialogs[LateBindingApi.Excel.Enums.XlBuiltInDialog.xlDialogApplyStyle].Show();
                    break;

                default:
                    throw (new Exception("Unkown dialog selected."));

            }

            string message = string.Format("The dialog returns {0}.", returnValue);
            MessageBox.Show(this, message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
  
         
            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

        }

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

    }
}
