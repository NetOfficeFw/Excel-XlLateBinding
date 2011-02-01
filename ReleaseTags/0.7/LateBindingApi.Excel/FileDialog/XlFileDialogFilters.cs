using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.FileDialog
{
    public class XlFileDialogFilters : XlNonCreatable , IEnumerable
    {
        #region Construction

        internal XlFileDialogFilters(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{		 

		}

        #endregion     
       
        #region Methods

        public XlFileDialogFilter Add(string description, string extensions)
        {
            object[] paramArray = new object[3];
            paramArray[0] = description;
            paramArray[1] = extensions;
            paramArray[2] = Missing.Value;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlFileDialogFilter newClass = new XlFileDialogFilter(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlFileDialogFilter Add(string description, string extensions, int position)
        {
      
            object[] paramArray = new object[3];
            paramArray[0] = description;
            paramArray[1] = extensions;
            paramArray[2] = position;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlFileDialogFilter newClass = new XlFileDialogFilter(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Clear()
        {
            InstanceType.InvokeMember("Clear", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Clear", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete(object filter)
        {
            object[] paramArray = new object[1];
            paramArray[0] = filter;
            InstanceType.InvokeMember("Clear", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlApplication Application
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlApplication newClass = new XlApplication(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        /// <summary>
        /// returns an XlFileDialogFilter by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlFileDialogFilter this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlFileDialogFilter newClass = new XlFileDialogFilter(this, returnValue);
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

        public int Creator 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
            XlFileDialogFilter[] res_addins = new XlFileDialogFilter[iCount];

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
