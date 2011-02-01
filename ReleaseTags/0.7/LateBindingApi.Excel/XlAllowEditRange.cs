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
    public class XlAllowEditRange: XlNonCreatable
    {
        #region Construction

        internal XlAllowEditRange(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{			 
		}

        #endregion

        #region Methods

        public void ChangePassword(string passWord)
        {
            object[] paramArray = new object[1];
            paramArray[0] = passWord;
            InstanceType.InvokeMember("ChangePassword", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Unprotect(string passWord)
        {
            object[] paramArray = new object[1];
            paramArray[0] = passWord;
            InstanceType.InvokeMember("Unprotect", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);

        }
        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlRange Range 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlRange newClass = new XlRange(this, returnValue);
                if (null == returnValue) return null; 
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlUserAccessList Users 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Users", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlUserAccessList newClass = new XlUserAccessList(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
         

        #endregion

        #region Scalar Properties

        public string Title
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Title", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Title", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }
}
