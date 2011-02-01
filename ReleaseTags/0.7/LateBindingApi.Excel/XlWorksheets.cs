using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel
{
    /// <summary>
    /// Represents the Worksheets in a Workbook as a Collection
    /// </summary>
    public class XlWorksheets  : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlWorksheets(IXlObject parentReference, object comReference) : base(parentReference, comReference)
        {

        }

        #endregion

        #region Foreach

        public IEnumerator GetEnumerator()
        {

            int iCount = Count;
            XlWorksheet[] res_sheets = new XlWorksheet[iCount];

            for (int i = 1; i <= iCount; i++)
                res_sheets[i - 1] = this[i];

            for (int i = 0; i < res_sheets.Length; i++)
            {
                yield return res_sheets[i];
            }

        }

        #endregion

        #region Scalar Properties

        /// <summary>
        /// Count of Worksheets
        /// </summary>
        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion

        #region ComReference Properties
        
        /// <summary>
        /// returns an Worksheet by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlWorksheet this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorksheet newClass = new XlWorksheet(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
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

        #endregion        
    }
}
