using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Web
{
    public class XlHTMLProject : XlNonCreatable
    {   
        #region Construction

        internal XlHTMLProject(IXlObject parentReference, object comReference) : base(parentReference, comReference)
		{           

		}

        #endregion

        #region Methods

        public void Open()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            InstanceType.InvokeMember("Open", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Open(MsoHTMLProjectOpen openKind)
        {
            object[] paramArray = new object[1];
            paramArray[0] = openKind;
            InstanceType.InvokeMember("Open", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RefreshDocument()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            InstanceType.InvokeMember("RefreshDocument", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RefreshProject()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            InstanceType.InvokeMember("RefreshProject", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlApplication Application
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlApplication newClass = new XlApplication(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlHTMLProjectItems HTMLProjectItems 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HTMLProjectItems", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlHTMLProjectItems newClass = new XlHTMLProjectItems(this, returnValue);
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

        public XlCreator Creator
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public MsoHTMLProjectState State 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("State", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoHTMLProjectState)returnValue;
            }
        }

        #endregion
    }
}
