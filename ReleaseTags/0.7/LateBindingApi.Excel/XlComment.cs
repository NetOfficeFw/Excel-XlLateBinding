using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel
{
    public class XlComment : XlNonCreatable
    {
        #region Construction

        internal XlComment(IXlObject parentReference, object comReference) : base(parentReference, comReference)
		{			 
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

        public XlApplication Shape 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Shape", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlApplication newClass = new XlApplication(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlNonCreatable Parent 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Parent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlNonCreatable newClass = XlDynamicType.CreateDynamicType(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public string Author 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Author", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public XlCreator Creator 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public bool Visible  
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region Methods

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlComment Next()
        {
            object returnValue = InstanceType.InvokeMember("Next", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            XlComment newClass = new XlComment(this, returnValue);
            if (null == returnValue) return null;
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlComment Previous()
        {
            object returnValue = InstanceType.InvokeMember("Previous", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            XlComment newClass = new XlComment(this, returnValue);
            if (null == returnValue) return null;
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public string Text()
        {
            object returnValue = InstanceType.InvokeMember("Text", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (string)returnValue;
        }
 
        public string Text(string text)
        {
            object[] paramArray = new object[3];
            paramArray[0] = text;
            paramArray[1] = Missing.Value;
            paramArray[2] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("Text", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (string)returnValue;
        }

        public string Text(string text, object start, object overwrite)
        {
            object[] paramArray = new object[3];
            paramArray[0] = text;
            paramArray[1] = start;
            paramArray[2] = overwrite;
            object returnValue = InstanceType.InvokeMember("Text", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (string)returnValue;
        }

        #endregion
    }
}
