﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Odbc
{
    public class XlODBCError : XlNonCreatable 
    {
        #region Construction

        internal XlODBCError(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{		 
		}

        #endregion       

        #region Methods
       
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

        public string ErrorString 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ErrorString", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string SqlState 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SqlState", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        #endregion
    }
}