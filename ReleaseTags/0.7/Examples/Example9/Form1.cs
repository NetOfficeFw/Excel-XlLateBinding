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

namespace Example9
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

            XlCommandBar commandBar;
            XlCommandBarPopup commandBarPop;
            XlCommandBarButton commandBarBtn;

            // first we enable the event support
            XlLateBindingApiSettings.EventsEnabled = true;
 
            // start excel and turn off msg boxes
            excelApplication = new XlApplication();
            excelApplication.DisplayAlerts = false;
         
            // add a new workbook
            XlWorkbook workBook = excelApplication.Workbooks.Add();

            #region Create a new menu
            
            // add a commandbar popup
            commandBarPop = (XlCommandBarPopup)excelApplication.CommandBars["Worksheet Menu Bar"].Controls.Add(
                                                           MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarPop.Caption = "commandBarPopup";
             
             // add a button to the popup
            #region a lot of words, how to access the picture
            /*
             you can see we use an own icon via .PasteFace()
             is not possible from outside process boundaries to use the PictureProperty directly
             the reason for is IPictureDisp: http://support.microsoft.com/kb/286460/de
             is not important is early or late binding or managed or unmanaged, the behaviour is always the same
             For example, a COMAddin running as InProcServer and can access the Picture Property
             Use the IconConverter.cs class from this project to convert a image to IPictureDisp
            */
            #endregion
            
            commandBarBtn = (XlCommandBarButton)commandBarPop.Controls.Add(MsoControlType.msoControlButton);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(this.Icon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.Click += new CommandBarButtonEvents_ClickEventHandler(commandBarBtn_Click);
            
            #endregion

            #region Create a new toolbar

            // add a new toolbar
            commandBar = excelApplication.CommandBars.Add("MyCommandBar", MsoBarPosition.msoBarTop, false, true);
            commandBar.Visible = true;

            // add a button to the toolbar
            commandBarBtn = (XlCommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            commandBarBtn.FaceId = 3;
            commandBarBtn.Click += new CommandBarButtonEvents_ClickEventHandler(commandBarBtn_Click);

            // add a dropdown box to the toolbar
            commandBarPop = (XlCommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup);
            commandBarPop.Caption = "commandBarPopup";

            // add a button to the popup, we use an own icon for the button
            commandBarBtn = (XlCommandBarButton)commandBarPop.Controls.Add(MsoControlType.msoControlButton);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            Clipboard.SetDataObject(this.Icon.ToBitmap());
            commandBarBtn.PasteFace();
            commandBarBtn.Click += new CommandBarButtonEvents_ClickEventHandler(commandBarBtn_Click);
            
            #endregion

            #region Create a new ContextMenu

            // add a commandbar popup
            commandBarPop = (XlCommandBarPopup)excelApplication.CommandBars["Cell"].Controls.Add(
                                                            MsoControlType.msoControlPopup, Missing.Value, Missing.Value, Missing.Value, true);
            commandBarPop.Caption = "commandBarPopup";

            // add a button to the popup
            commandBarBtn = (XlCommandBarButton)commandBarPop.Controls.Add(MsoControlType.msoControlButton);
            commandBarBtn.Style = MsoButtonStyle.msoButtonIconAndCaption;
            commandBarBtn.Caption = "commandBarButton";
            commandBarBtn.FaceId = 9;
            commandBarBtn.Click += new CommandBarButtonEvents_ClickEventHandler(commandBarBtn_Click);

            #endregion

            #region Display info
           
            workBook.Worksheets[1].Cells(2, 2).Value = "this excel instance contains 3 custom menus"; 
            workBook.Worksheets[1].Cells(3, 2).Value = "the main menu, the toolbar menu and the cell context menu";
            workBook.Worksheets[1].Cells(4, 2).Value = "in this case the menus are temporaily created";
            workBook.Worksheets[1].Cells(5, 2).Value = "they are not persistant and needs no unload event or something like this";
            workBook.Worksheets[1].Cells(6, 2).Value = "you can also create persistant menus if you want";

            #endregion

            excelApplication.Visible = true;

            button1.Enabled = false;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            excelApplication.Quit();
            excelApplication.Dispose();

            button1.Enabled = true;
            button2.Enabled = false;         
        }

        void commandBarBtn_Click(XlCommandBarButton Ctrl, ref bool CancelDefault)
        {
            textBoxEvents.BeginInvoke(_updateDelegate, new object[] { "Click called." });
        }

        private void UpdateTextbox(string Message)
        {
            textBoxEvents.AppendText(Message + "\r\n");
        }

    }
}
