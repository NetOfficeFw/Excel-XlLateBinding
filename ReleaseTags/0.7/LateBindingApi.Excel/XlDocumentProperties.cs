﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel
{
    public class XlDocumentProperties : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlDocumentProperties(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion
        
        #region Methods

        public XlDocumentProperty Add(string name, object value)
        {
            object[] paramArray = new object[1];
            paramArray[0] = name;
            paramArray[1] = value;
            object returnValue = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlDocumentProperty newClass = new XlDocumentProperty(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
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

        /// <summary>
        /// returns an XlDocumentProperty by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlDocumentProperty this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDocumentProperty newClass = new XlDocumentProperty(this, returnValue);
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

        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion
        
        #region Foreach

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            int iCount = Count;
            XlDocumentProperty[] res_addins = new XlDocumentProperty[iCount];

            for (int i = 1; i <= iCount; i++)
                res_addins[i - 1] = this[i];

            for (int i = 0; i < res_addins.Length; i++)
            {
                yield return res_addins[i];
            }

        }

        #endregion
    }
}
