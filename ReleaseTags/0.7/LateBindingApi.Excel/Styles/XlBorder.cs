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
    /// Represents a Border from an ExcelRange
    /// </summary>
    public class XlBorder : XlNonCreatable
    {        
        #region Construction

        internal XlBorder(IXlObject parentReference, object comReference):base(parentReference,comReference)
		{
			 
		}

        #endregion
      
        #region Scalar Properties

        /// <summary>
        /// Color of Border Line
        /// </summary>
        public XlColorIndex  ColorIndex
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ColorIndex", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlColorIndex)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ColorIndex", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        /// <summary>
        /// Color of Border Line
        /// </summary>
        public double Color
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Color", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Color", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        /// <summary>
        /// Weight of Border Line
        /// </summary>
        public int Weight
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Weight", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Weight", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        /// <summary>
        /// LineStyle of Border Line
        /// </summary>
        public XlLineStyle LineStyle 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LineStyle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlLineStyle)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("LineStyle", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }         
        }

        #endregion     
    }
}
