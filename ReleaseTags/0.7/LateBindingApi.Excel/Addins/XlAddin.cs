using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Addins
{
    /// <summary>
    /// Represents an Addin in Excel
    /// </summary>
    public class XlAddin : XlNonCreatable
    {
        #region Construction

        internal XlAddin(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
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

        public string CLSID
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CLSID", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string FullName
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FullName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string Path
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Path", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public bool Installed
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Installed", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        #endregion
    }
}
