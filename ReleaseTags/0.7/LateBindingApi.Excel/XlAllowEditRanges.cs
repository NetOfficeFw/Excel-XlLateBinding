using System;
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
    public class XlAllowEditRanges : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlAllowEditRanges(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion
        
        #region Methods       

        public XlAllowEditRange Add(string title, XlRange range)
        {
            object[] paramArray = new object[1];
            paramArray[0] = title;
            paramArray[1] = range;
            object returnValue = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlAllowEditRange newClass = new XlAllowEditRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        #endregion
        
        #region COMReference Properties

        /// <summary>
        /// returns an XlAllowEditRange by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlAllowEditRange this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlAllowEditRange newClass = new XlAllowEditRange(this, returnValue);
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
            XlAllowEditRange[] res_addins = new XlAllowEditRange[iCount];

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
