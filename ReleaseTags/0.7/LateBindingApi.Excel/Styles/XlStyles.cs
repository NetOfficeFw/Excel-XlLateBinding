using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Styles
{
    public class XlStyles : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Properties

        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public XlStyle this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object comRef  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                XlStyle newClass = new XlStyle(this, comRef);
                ListChildReferences.Add(newClass);
                return newClass;
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
            XlStyle[] res_addins = new XlStyle[iCount];

            for (int i = 1; i <= iCount; i++)
                res_addins[i - 1] = this[i];

            for (int i = 0; i < res_addins.Length; i++)
            {
                yield return res_addins[i];
            }

        }

        #endregion

        #region Construction

        /// <summary>
        /// IXlNonCreatable Constructor
        /// </summary>
        /// <param name="parentReference"></param>
        /// <param name="comReference"></param>
        internal XlStyles(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{			 
		}

        #endregion  
        
        #region Methods

        public XlStyle Add(string name)
        {
            object[] paramArray = new object[2];
            paramArray[0] = name;
            paramArray[1] = Missing.Value;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            XlStyle newClass = new XlStyle(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlStyle Add(string name, string basedOn)
        {
            object[] paramArray = new object[2];
            paramArray[0] = name;
            paramArray[1] = basedOn;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            XlStyle newClass = new XlStyle(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        #endregion
    }
}
