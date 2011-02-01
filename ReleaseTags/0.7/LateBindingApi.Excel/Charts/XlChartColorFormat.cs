using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Charts
{
    public class XlChartColorFormat : XlNonCreatable
    {   
        #region Construction

        internal XlChartColorFormat(IXlObject parentReference, object comReference) : base(parentReference, comReference)
		{			 
		}

        #endregion  

        #region Scalar Properties

        public int RGB
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RGB", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int SchemeColor
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SchemeColor", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("SchemeColor", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Type
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion
    }
}
