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
    public class XlSheets : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlSheets(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion  

        #region Methods

        public XlWorksheet Add()
        {
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWorksheet newClass = new XlWorksheet(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Copy()
        {
            InstanceType.InvokeMember("Copy", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);           
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void FillAcrossSheets(XlRange range, XlFillWith type)
        {
            object[] paramArray = new object[2];
            paramArray[0] = range.COMReference;
            paramArray[1] = type;
            InstanceType.InvokeMember("FillAcrossSheets", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Move()
        {
            InstanceType.InvokeMember("Move", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PrintOut()
        {
            InstanceType.InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PrintPreview()
        {
            InstanceType.InvokeMember("PrintPreview", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlHPageBreaks HPageBreaks
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HPageBreaks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlHPageBreaks newClass = new XlHPageBreaks(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVPageBreaks VPageBreaks
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VPageBreaks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVPageBreaks newClass = new XlVPageBreaks(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWorksheet this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorksheet newClass = new XlWorksheet(this, returnValue);
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
            XlWorksheet[] res_addins = new XlWorksheet[iCount];

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
