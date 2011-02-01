using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Core;
using Excel = LateBindingApi.Excel;

namespace Tutorial01
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
            Excel.Application application = new LateBindingApi.Excel.Application();
            application.DisplayAlerts = false;
            /*
            *  LateBindingApi.Excel manages COM Proxies for you to avoid any kind of memory leaks
            *  and make sure your excel instance removes from process list if you want.
            */
            
            Excel.Workbook book = application.Workbooks.Add();
            /* 
            * now we have 2 new COM Proxies created.
            * 
            * the first proxy was created while accessing the Workbooks collection from application
            * the second proxy was created by the Add() method from Workbooks and stored now in book
            */


            Excel.Worksheet sheet = book.Worksheets.Add();
            /*
            * we create also 2 new COM Proxies.
            * 
            * the first proxy was created while accessing the Worksheets collection from book
            * the second proxy was created by the Add() method from Worksheets and stored now in sheet
            * 
            * the 2 proxies was created about the book object and are managed now by the book object
            * they are now 'child proxies' collected in a internal list.
            */


            book.Dispose();
            /* now we have disposed the Workbook, the inner Proxy will be released */
            /* all child proxies there was created about the book are also released now */
            /* this means the 2 proxies there we have created while add a new worksheet */
            /* the local stored variable Excel.Worksheet sheet are not longer valid */
 

            /*
            * SUMMARY: 
            * With LateBindingApi.Excel you have the choice to use Dispose() for any object or doesnt.
            * In case of doesn't dont forget to dispose the application and all child proxies will be released automaticly. 
            * It is necessary to understand the demonstrated child proxy concept of LateBindingApi.Excel to avoid any suprises. 
            * With these concept you have a powerfull third way. Use child proxies without Dispose() to get more readable & smaller code.
            * Dispose the parent object of your childs and thats all. 
            */


            /* another example for efficient using child proxy management */
            book = application.Workbooks.Add();            
            foreach (Excel.Worksheet item in book.Worksheets)
            {
                for (int i = 1; i <= 100; i++)
                    item.Cells[1, i].Value = "Hello World";             
            }
            book.Dispose();
            /*
            * we create a new workbook with default 3 new worksheets 
            * we do a foreach to all worksheets and write 100 times "Hello World" in cells
            * we have created now 300 range proxies 3 worksheet proxies 1 worksheet enumerator proxy and the book proxy
            * the parent object is ouer book, diposed it and the child proxies will be also disposed
            * Note: its not allowed in the MS Interop Assemblies to use an enumerator like these.
            * you must store it in a local variable first and dispose explicitly after using.
            * With LateBindingApi.Excel you can do this without any problems. 
            */


            application.Quit();
            application.Dispose();
            /*the application object is ouer root object*/
            /*dispose them will release himself and any childs*/
            /*the excel instance are now removed from process list*/

            MessageBox.Show(this, "Done!",this.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);   
        }
    }
}
