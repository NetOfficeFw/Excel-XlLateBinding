using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = LateBindingApi.Excel;
using Office = LateBindingApi.Office;

namespace Tutorial02
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

            /*
            *  COMVariant is a LateBindingApi defined Type as substitute for the COM Datatype Variant. 
            *
            *  To work with COMVariant use there following properties:
            *  object UnderlyingObject - the mapped original object
            *  bool IsCOMProxy         - UnderlyingObject is a COM Proxy or native type
            *  string TypeName         - the name of UnderlyingObject type
            *
            *  You can cast the COMVariant directly in another LateBindingApi Type, see the following example
            */

            /* Example: */
            /* the property Selection from Excel.Application is defined as COMVariant */
            /* It can have multiple types at runtime what is currently selected, a worksheet, a range, a window or other*/

            // select worksheet for example
            application.Workbooks[1].Worksheets[1].Select();

            COMVariant myVariant = application.Selection;
            if (null != myVariant)
            {
                switch (myVariant.TypeName)
                {
                    case "Worksheet":
                        Excel.Worksheet sheet = (Excel.Worksheet)myVariant;
                        break;

                    case "Range":
                        Excel.Range range = (Excel.Range)myVariant;
                        break;
                } 
            }


            /* Another example: */
            /* GetOpenFileName returns a COMVariant there can be a string or a boolean in case of user clicks cancel */
            myVariant = application.GetOpenFilename("Text Files (*.txt), *.txt");
            if (null != myVariant)
            {
                string name = myVariant.TypeName;
                object underlyingObject = myVariant.UnderlyingObject;

                string message = string.Format("GetOpenFilename returns a {0}\r\n{1}", name, underlyingObject);
                MessageBox.Show(this, message, this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information); 
            }

            /* Last example: */
            /* a lot of scalar properties defined as variant in excel (the reason is still unkown)*/
            /* use the multiple overload helper function ToCOMVariant */
            /* Note: a other way to handle multiple types in a variable is the Invoker, see Tutorial03 */
            Excel.Range sampleRange = application.Workbooks[1].Worksheets[1].Cells[1, 1];                       
            int colorIndex = (int)sampleRange.Font.ColorIndex.UnderlyingObject;
            sampleRange.Font.ColorIndex = COMVariant.ToCOMVariant(colorIndex + 1);

            /*
            *  COMObject inherited COMVariant and is a LateBindingApi defined Type as substitute for a unkown Dispatch Interface aka System._ComProxy in .NET.
            *  All LateBindingApi WrapperClasses inherites COMObject 
            *  You can cast COMObject directly in another LateBindingApi Type, see the following example
            *  To work with COMObject use the following properties:
            *  object UnderlyingObject - the mapped original System._ComProxy object
            *  string TypeName         - the name of UnderlyingObject type
            */
            foreach (Office.COMAddIn item in application.COMAddIns)
            {
                /* COMAddIn.Application is defined as COMObject*/
                string name = item.Application.TypeName;              
                Excel.Application parentApplication = (Excel.Application)item.Application;
            }

            application.Quit();
            application.Dispose();
        }
    }
}
