using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlAddin : XlNonCreatable
    {
        #region Construction

        internal XlAddin(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
		 
		}

        #endregion  

        #region COMReference
        
        public XlAddins Collection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Collection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlAddins newClass = new XlAddins(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlVBE VBE
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBE", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBE newClass = new XlVBE(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Properties

        public bool Connect 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Connect", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public string Description 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Description", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Description", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        public object Object 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Object", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
        }

        public string ProgId 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProgId", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        #endregion
    }
}
