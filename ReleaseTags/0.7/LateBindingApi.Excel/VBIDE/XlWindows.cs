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
    public class XlWindows : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlWindows(IXlObject parentReference, object comReference) : base(parentReference, comReference)
        {
        }

        #endregion  

        #region COMReference Properties
       
        public XlVBE VBE
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBE", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlVBE newClass = new XlVBE(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        /// <summary>
        /// returns an XlWindow by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlWindow this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWindow newClass = new XlWindow(this, returnValue);
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
        
        #region Methods

        public XlWindow CreateToolWindow(XlAddin addInInst, string progId, string caption, string guidPosition, object docObj)
        {
            object[] paramArray = new object[5];
            paramArray[0] = addInInst.COMReference;
            paramArray[1] = progId;
            paramArray[2] = caption;
            paramArray[3] = guidPosition;
            paramArray[4] = docObj;
            object returnValue  = InstanceType.InvokeMember("CreateToolWindow", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWindow newClass = new XlWindow(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
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
            XlWindow[] res_addins = new XlWindow[iCount];

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
