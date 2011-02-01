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
    public class XlReferences : XlNonCreatable, System.Collections.IEnumerable
    { 
        #region Construction

        internal XlReferences(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
		 
		}

        #endregion

        #region COMReference Properties

        /// <summary>
        /// returns an XlReference by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlReference this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlReference newClass = new XlReference(this, returnValue);
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

        public XlVBProject Parent
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Parent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBProject newClass = new XlVBProject(this, returnValue);
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
            XlReference[] res_addins = new XlReference[iCount];

            for (int i = 1; i <= iCount; i++)
                res_addins[i - 1] = this[i];

            for (int i = 0; i < res_addins.Length; i++)
            {
                yield return res_addins[i];
            }

        }

        #endregion

        #region Methods

        public XlReference AddFromFile(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            object returnValue  = InstanceType.InvokeMember("AddFromFile", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlReference newClass = new XlReference(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlReference AddFromGuid(string guid, int major, int minor)
        {
            object[] paramArray = new object[3];
            paramArray[0] = guid;
            paramArray[1] = major;
            paramArray[2] = minor;
            object returnValue  = InstanceType.InvokeMember("AddFromGuid", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlReference newClass = new XlReference(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Remove(XlReference reference)
        {
            object[] paramArray = new object[1];
            paramArray[0] = reference.COMReference;
            InstanceType.InvokeMember("Remove", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion
    }
}
