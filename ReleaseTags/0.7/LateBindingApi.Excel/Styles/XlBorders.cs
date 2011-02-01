using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Styles
{
    /// <summary>
    /// Represents all Borders from an ExcelRange as a Collection
    /// </summary>
    public class XlBorders : XlNonCreatable
    {              
        #region Construction

        internal XlBorders(IXlObject parentReference, object comReference):base(parentReference,comReference)
		{
		 
		}

        #endregion
        
        #region ComReference Properties

        /// <summary>
        /// Gets a Border Object by XlBordersIndex 
        /// </summary>
        /// <param name="BordersIndex"></param>
        /// <returns></returns>
        public XlBorder this[XlBordersIndex bordersIndex]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = bordersIndex;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty| BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlBorder newClass = new XlBorder(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;         
            }
        }

        #endregion 
    }

}
