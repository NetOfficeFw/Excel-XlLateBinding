using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Office
{
    public class XlMsoEnvelope : XlNonCreatable 
    {
        #region Construction

        internal XlMsoEnvelope(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
        
		}

        #endregion

        #region Scalar Properties

        public string Introduction
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Introduction", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Introduction", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion   
    }
}
