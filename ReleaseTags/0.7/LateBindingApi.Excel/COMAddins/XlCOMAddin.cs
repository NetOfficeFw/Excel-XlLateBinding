using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.COMAddins
{
    /// <summary>
    /// Represents a COMAddin in Excel
    /// </summary>
    public class XlCOMAddin : XlNonCreatable 
    {  
        #region Construction

        internal XlCOMAddin(IXlObject parentReference, object comReference):base(parentReference,comReference)
		{		 
		}

        #endregion   
       
        #region Methods
        
        /// <summary>
        /// Given Object from Addin, can be NULL, must be released by Caller with Marshal.ReleaseComObject!
        /// </summary>
        public object Object()
        {
            object returnValue  = InstanceType.InvokeMember("Object", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return returnValue;
        }

        #endregion

        #region Scalar Properties

        public string ProgId
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProgId", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string Guid
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Guid", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string Description
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Description", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public bool Connect
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Connect", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        #endregion
    }
}
