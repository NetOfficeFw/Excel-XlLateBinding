using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Shapes
{
    public class XlGroupShapes : XlNonCreatable, System.Collections.IEnumerable
    {   
        #region Construction

        internal XlGroupShapes(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion

        #region Methods

        public XlShapeRange Range(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue  = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;  
            XlShapeRange newClass = new XlShapeRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        #endregion

        #region COMReference Properties

        /// <summary>
        /// returns an XlGroupShape by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlShape this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShape newClass = new XlShape(this, returnValue);
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
            XlShape[] res_addins = new XlShape[iCount];

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
