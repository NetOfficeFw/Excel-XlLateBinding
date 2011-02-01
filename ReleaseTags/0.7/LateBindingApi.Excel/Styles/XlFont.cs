using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Styles
{
    public class XlFont : XlNonCreatable
    {        
        #region Construction

        internal XlFont(IXlObject parentReference, object comReference):base(parentReference,comReference)
		{
		 
		}

        #endregion

        #region Scalar Properties

        /// <summary>
        /// Name of Font 
        /// </summary>
        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }                             
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Name", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);                 
            }
        }

        /// <summary>
        /// ForeColor of Font 
        /// </summary>
        public double Color
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Color", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Color", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        /// <summary>
        /// Size of Font
        /// </summary>
        public double Size
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Size", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Size", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        /// <summary>
        /// Bold of Font
        /// </summary>
        public bool Bold
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Bold", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Bold", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        /// <summary>
        /// Underline of Font
        /// </summary>
        public bool Underline
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Underline", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Underline", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        /// <summary>
        /// Italic of Font
        /// </summary>
        public bool Italic
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Italic", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Italic", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }
}
