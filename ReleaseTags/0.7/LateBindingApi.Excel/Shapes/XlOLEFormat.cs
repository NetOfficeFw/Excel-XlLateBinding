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
    public class XlOLEFormat : XlNonCreatable
    {
        #region Construction

        internal XlOLEFormat(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion

        #region Methods

        public void Activate()
        {
            InstanceType.InvokeMember("Activate", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region Scalar Properties

        public string progID
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("progID", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public object Object
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Object", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
        }

        #endregion
    }
}
