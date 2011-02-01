using System;
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
    public class XlVBProjects : XlNonCreatable, System.Collections.IEnumerable
   {
        #region Construction

        internal XlVBProjects(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{			 
		}

        #endregion  

        #region COMReference Properties
        
        /// <summary>
        /// returns an XlVBProject by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlVBProject this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        #region Methods

        public void Remove(XlVBProject lpc)
        {
            object[] parameter = new object[1];
            parameter[0] = lpc;
            InstanceType.InvokeMember("Remove", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlVBProject Add(vbext_ProjectType type)
        {
            object[] parameter = new object[1];
            parameter[0] = type;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlVBProject newClass = new XlVBProject(this, returnValue);
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
            XlVBProject[] res_addins = new XlVBProject[iCount];

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
