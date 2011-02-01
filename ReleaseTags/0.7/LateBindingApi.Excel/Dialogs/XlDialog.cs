using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Dialogs
{
    public class XlDialog : XlNonCreatable 
    {
        #region Construction

        internal XlDialog(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{		 

		}

        #endregion     

        #region Methods

        public bool Show()
        {
            object returnValue  = InstanceType.InvokeMember("Show", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        #endregion

    }
}
