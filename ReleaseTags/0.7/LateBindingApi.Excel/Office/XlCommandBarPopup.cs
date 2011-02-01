using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Office
{
    public class XlCommandBarPopup : XlCommandBarControl
    {
        #region Construction

        internal XlCommandBarPopup(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{			 
		}

        #endregion 

        #region COMReference Properties

        public XlCommandBarControls Controls
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Controls", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBarControls newClass = new XlCommandBarControls(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion
    }
}
