using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Pivot
{
    public class XlPivotLayout : XlNonCreatable
    {  
        #region Construction

        internal XlPivotLayout(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{			 
		}

        #endregion 

        #region COMReference Properties

        public XlPivotTable PivotTable
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PivotTable", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPivotTable newClass = new XlPivotTable(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion
    }
}
