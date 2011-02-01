using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

using LateBindingApi.Core;
using Excel = LateBindingApi.Excel;
using LateBindingApi.Excel.Enums;

namespace Example8
{
    public partial class Form1 : Form
    {
        Excel.Application _excelApplication;

        private delegate void UpdateEventTextDelegate(string Message);
        UpdateEventTextDelegate _updateDelegate;

        public Form1()
        {
            InitializeComponent();
            _updateDelegate = new UpdateEventTextDelegate(UpdateTextbox);

            /*
            * Initialize Api COMObject & COMVariant Support
            */
            LateBindingApi.Core.Factory.Initialize();
        }
     
        private void button1_Click(object sender, EventArgs e)
        {
            // we enable the event support
            LateBindingApi.Core.Settings.EnableEvents = true;
 
            // start excel and turn off msg boxes
            _excelApplication = new Excel.Application();
            _excelApplication.DisplayAlerts = false;
            _excelApplication.Visible = true;

            /*
            we register some events. note: the event trigger was called from excel, means an other Thread
            remove the Quit() call below and check out more events if you want
            you can get event notifys from various objects: Application or Workbook or Worksheet for example
            */
            _excelApplication.NewWorkbookEvent += new Excel.AppEvents_NewWorkbookEventHandler(excelApplication_NewWorkbook);
            _excelApplication.WorkbookBeforeCloseEvent += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(excelApplication_WorkbookBeforeClose);
            _excelApplication.WorkbookActivateEvent += new Excel.AppEvents_WorkbookActivateEventHandler(excelApplication_WorkbookActivate);
            _excelApplication.WorkbookDeactivateEvent += new Excel.AppEvents_WorkbookDeactivateEventHandler(excelApplication_WorkbookDeactivate);
            _excelApplication.SheetActivateEvent += new Excel.AppEvents_SheetActivateEventHandler(excelApplication_SheetActivate);
            _excelApplication.SheetDeactivateEvent += new Excel.AppEvents_SheetDeactivateEventHandler(excelApplication_SheetDeactivate);

            // add a new workbook add a sheet and close
            Excel.Workbook workBook = _excelApplication.Workbooks.Add();
            workBook.Worksheets.Add();
            workBook.Close();

            _excelApplication.Quit();
            _excelApplication.Dispose();
        }

        void excelApplication_NewWorkbook(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event NewWorkbook called." });
        }

        void excelApplication_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookBeforeClose called." });
        }

        void excelApplication_WorkbookActivate(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookActivate called." });
        }

        void excelApplication_WorkbookDeactivate(Excel.Workbook Wb)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event WorkbookDeactivate called." });
        }

        void excelApplication_SheetActivate(COMObject Sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetActivate called." });
        }

        void excelApplication_SheetDeactivate(COMObject Sh)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Event SheetDeactivate called." });
        }

        private void UpdateTextbox(string Message)
        {
            textBoxEvents.AppendText(Message+"\r\n");
        }

    }
}
    