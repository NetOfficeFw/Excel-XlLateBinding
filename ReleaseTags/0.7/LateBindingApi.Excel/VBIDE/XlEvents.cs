using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlEvents : XlNonCreatable
    {
        #region Construction

        internal XlEvents(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion

        #region Methods

        public XlCommandBarEvents CommandBarEvents(object CommandBarControl)
        {
            object[] paramArray = new object[1];
            paramArray[0] = CommandBarControl;
            object returnValue  = InstanceType.InvokeMember("CommandBarEvents", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCommandBarEvents newClass = new XlCommandBarEvents(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;          
        }

        public XlReferencesEvents ReferencesEvents(XlVBProject VBProject)
        {
            object[] paramArray = new object[1];
            paramArray[0] = VBProject;
            object returnValue  = InstanceType.InvokeMember("ReferencesEvents", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlReferencesEvents newClass = new XlReferencesEvents(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }
        
        #endregion
    }
}
