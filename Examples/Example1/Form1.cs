using System;
using System.Drawing;
using System.Windows.Forms;

using Excel = LateBindingApi.Excel;
using LateBindingApi.Excel.Enums; 

namespace Example1
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

            /*do background color for cells*/
           
            string listSeperator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            
            // draw the face
            string rangeAdressFace = string.Format("$C10:$M10{0}$C30:$M30{0}$C11:$C30{0}$M11:$M30", listSeperator);
            workSheet.get_Range(rangeAdressFace).Interior.Color = ToDouble(Color.DarkGreen);
         
            string rangeAdressEyes = string.Format("$F14{0}$J14", listSeperator);
            workSheet.get_Range(rangeAdressEyes).Interior.Color = ToDouble(Color.Black);

            string rangeAdressNoise = string.Format("$G18:$I19", listSeperator);
            workSheet.get_Range(rangeAdressNoise).Interior.Color = ToDouble(Color.DarkGreen);

            string rangeAdressMouth = string.Format("$F26{0}$J26{0}$G27:$I27", listSeperator);
            workSheet.get_Range(rangeAdressMouth).Interior.Color = ToDouble(Color.DarkGreen);
                 
            // border the face with the border arround method
            workSheet.get_Range(rangeAdressFace).BorderAround(XlLineStyle.xlDashDot, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexNone, Color.BlueViolet.ToArgb());
            workSheet.get_Range(rangeAdressEyes).BorderAround(XlLineStyle.xlDashDot, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexNone, Color.BlueViolet.ToArgb());
            workSheet.get_Range(rangeAdressNoise).BorderAround(XlLineStyle.xlDouble, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexNone, Color.BlueViolet.ToArgb());
            
            // border explicitly
            workSheet.get_Range(rangeAdressMouth).Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
            workSheet.get_Range(rangeAdressMouth).Borders[XlBordersIndex.xlEdgeBottom].Weight = 4;
            workSheet.get_Range(rangeAdressMouth).Borders[XlBordersIndex.xlEdgeBottom].Color = ToDouble(Color.Black);

            // save the book 
            string fileExtension = GetDefaultExtension(_excelApplication);
            string workbookFile = string.Format("{0}\\Example01{1}",  Environment.CurrentDirectory,fileExtension);
            workBook.SaveAs(workbookFile);
            
            // close excel and dispose reference
            _excelApplication.Quit();
            _excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
        }
        
        #region Helper

        /// <summary>
        /// Translate a color to double
        /// </summary>
        /// <param name="color">expression to convert</param>
        /// <returns>color</returns>
        private static double ToDouble(System.Drawing.Color color)
        {
            uint returnValue = color.B;
            returnValue = returnValue << 8;
            returnValue += color.G;
            returnValue = returnValue << 8;
            returnValue += color.R;
            return returnValue;
        }

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
