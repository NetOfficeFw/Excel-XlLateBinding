using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Views
{
    public class XlCustomView : XlNonCreatable 
    {
        #region Construction

        internal XlCustomView(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion 

        #region Methods

        public bool Delete()
        {
            object returnValue  = InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public bool Show()
        {
            object returnValue  = InstanceType.InvokeMember("Show", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        #endregion

        #region COMReference Properties
        
        public XlApplication Application
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlApplication newClass = new XlApplication(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        #endregion

        #region Scalar Properties

        public XlCreator Creator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public bool PrintSettings
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PrintSettings", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool RowColSettings
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RowColSettings", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        #endregion 
    }
}
