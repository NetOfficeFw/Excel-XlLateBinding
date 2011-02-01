using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = LateBindingApi.Excel;
using Office = LateBindingApi.Office;

namespace Tutorial03
{
    public partial class Form1 : Form
    {
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
            Excel.Application application = new Excel.Application();
            application.DisplayAlerts = false;
            application.Workbooks.Add();
           
            Excel.Range sampleRange = application.Workbooks[1].Worksheets[1].Cells[1, 1];

            // we set the COMVariant ColorIndex from Font of ouer sample range with the invoker class
            // sometimes its easier than using COMVariant.ToCOMVariant Method or no overload for the underlying target type exists
            Invoker.PropertySet(sampleRange.Font, "ColorIndex", 1);

            // creates a native unmanaged ComProxy
            object comProxy = Invoker.MethodReturn(application, "Workbooks");
            Marshal.ReleaseComObject(comProxy); 

            application.Quit();
            application.Dispose();
        }
    }
}
