using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Office;
using LateBindingApi.Excel.Enums;

namespace Example8
{

    public partial class Form1 : Form
    {
        XlApplication excelApplication;

        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Form1()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);
        }
     
        private void button1_Click(object sender, EventArgs e)
        {
            // we enable the event support
            XlLateBindingApiSettings.EventsEnabled = true;     
            
            // start excel and turn off msg boxes
            excelApplication = new XlApplication();
            excelApplication.DisplayAlerts = false;
            excelApplication.Visible = true;

            /*
            we register some events. note: the event trigger was called from excel, means an other Thread
            remove the Quit() call below and check out more events if you want
            you can get event notifys from various objects: XlApplication or XlWorkbook or XlWorksheet for example
            */
            excelApplication.NewWorkbook += new AppEvents_NewWorkbookEventHandler(excelApplication_NewWorkbook);
            excelApplication.WorkbookBeforeClose += new AppEvents_WorkbookBeforeCloseEventHandler(excelApplication_WorkbookBeforeClose);
            excelApplication.WorkbookActivate += new AppEvents_WorkbookActivateEventHandler(excelApplication_WorkbookActivate);
            excelApplication.WorkbookDeactivate += new AppEvents_WorkbookDeactivateEventHandler(excelApplication_WorkbookDeactivate);
            excelApplication.SheetActivate +=new AppEvents_SheetActivateEventHandler(excelApplication_SheetActivate);
            excelApplication.SheetDeactivate += new AppEvents_SheetDeactivateEventHandler(excelApplication_SheetDeactivate);
           
            // add a new workbook add a sheet and close
            XlWorkbook workBook = excelApplication.Workbooks.Add();
            workBook.Worksheets.Add();
            workBook.Close();

            excelApplication.Quit();
            excelApplication.Dispose();
        }

      
        void excelApplication_NewWorkbook(XlWorkbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event NewWorkbook called." });
        }

        void excelApplication_WorkbookBeforeClose(XlWorkbook Wb, ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookBeforeClose called." });
        }

        void excelApplication_WorkbookActivate(XlWorkbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookActivate called." });
        }

        void excelApplication_WorkbookDeactivate(XlWorkbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookDeactivate called." });
        }       
      
        void excelApplication_SheetActivate(XlWorksheet Sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetActivate called." });
        }

        void excelApplication_SheetDeactivate(XlWorksheet Sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetDeactivate called." });
        }

        private void UpdateTextbox(string Message)
        {
            textBoxEvents.AppendText(Message+"\r\n");
        }

    }
}
    