using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Enums; 

namespace Example1
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
            
            /*do background color for cells*/

            string listSeperator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            
            // draw the face
            string rangeAdressFace = string.Format("$C10:$M10{0}$C30:$M30{0}$C11:$C30{0}$M11:$M30", listSeperator);
            workSheet.Range(rangeAdressFace).Interior.Color = XlConverter.ToDouble(Color.DarkGreen);
           
            string rangeAdressEyes = string.Format("$F14{0}$J14", listSeperator);
            workSheet.Range(rangeAdressEyes).Interior.Color = XlConverter.ToDouble(Color.Black);

            string rangeAdressNoise = string.Format("$G18:$I19", listSeperator);
            workSheet.Range(rangeAdressNoise).Interior.Color = XlConverter.ToDouble(Color.DarkGreen);

            string rangeAdressMouth = string.Format("$F26{0}$J26{0}$G27:$I27", listSeperator);
            workSheet.Range(rangeAdressMouth).Interior.Color = XlConverter.ToDouble(Color.DarkGreen);
    

            /*do borderlines for cells*/
             
            // border the face with the border arround method
            workSheet.Range(rangeAdressFace).BorderAround(LateBindingApi.Excel.Enums.XlLineStyle.xlDashDot, LateBindingApi.Excel.Enums.XlBorderWeight.xlThin, Color.BlueViolet.ToArgb());
            workSheet.Range(rangeAdressEyes).BorderAround(LateBindingApi.Excel.Enums.XlLineStyle.xlDashDot, LateBindingApi.Excel.Enums.XlBorderWeight.xlThin, Color.BlueViolet.ToArgb());
            workSheet.Range(rangeAdressNoise).BorderAround(LateBindingApi.Excel.Enums.XlLineStyle.xlDouble, LateBindingApi.Excel.Enums.XlBorderWeight.xlThin, Color.BlueViolet.ToArgb());
            
            // border explicitly
            workSheet.Range(rangeAdressMouth).Borders[LateBindingApi.Excel.Enums.XlBordersIndex.xlEdgeBottom].LineStyle = LateBindingApi.Excel.Enums.XlLineStyle.xlDouble;
            workSheet.Range(rangeAdressMouth).Borders[LateBindingApi.Excel.Enums.XlBordersIndex.xlEdgeBottom].Weight = 4;
            workSheet.Range(rangeAdressMouth).Borders[LateBindingApi.Excel.Enums.XlBordersIndex.xlEdgeBottom].Color = 400;

            // save the book 
            string fileExtension = XlConverter.GetDefaultExtension(excelApplication);
            string workbookFile = string.Format("{0}\\Example1{1}",  Environment.CurrentDirectory,fileExtension);
            workBook.SaveAs(workbookFile);
            
            // close excel and dispose reference
            excelApplication.Quit();
            excelApplication.Dispose();

            FinishDialog fDialog = new FinishDialog("Workbook saved.", workbookFile);
            fDialog.ShowDialog(this);
            
        }
        
       
    }
}
