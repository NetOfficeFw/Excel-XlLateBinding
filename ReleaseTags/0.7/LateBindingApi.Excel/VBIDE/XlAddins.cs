﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlAddins : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlAddins(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion 

        #region Methods

        public void Update()
        {
            InstanceType.InvokeMember("Update", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }
        
        #endregion

        #region COMReference Properties

        /// <summary>
        /// returns an XlAddin by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlAddin this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlAddin newClass = new XlAddin(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBE VBE
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBE", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBE newClass = new XlVBE(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }


        #endregion

        #region Scalar Properties

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
            XlAddin[] res_addins = new XlAddin[iCount];

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
