using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Office;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlVBProject : XlNonCreatable
    {      
        #region Construction

        internal XlVBProject(IXlObject parentReference, object comReference) : base(parentReference, comReference)
        {
		 
		}

        #endregion

        #region Methods

        public void MakeCompiledFile()
        {
            InstanceType.InvokeMember("MakeCompiledFile", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlVBProjects Collection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Collection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBProjects newClass = new XlVBProjects(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlReferences References
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("References", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlReferences newClass = new XlReferences(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBComponents VBComponents
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBComponents", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBComponents newClass = new XlVBComponents(this, returnValue);
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

        #region Scalar Properties

        public string BuildFileName 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BuildFileName", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("BuildFileName", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        public string FileName 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FileName", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public int HelpContextID 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HelpContextID", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HelpContextID", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string HelpFile 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HelpFile", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HelpFile", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

 
        public vbext_VBAMode Mode 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Mode", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (vbext_VBAMode)returnValue;
            }
        }

        public string Name 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public vbext_ProjectProtection Protection 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Protection", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (vbext_ProjectProtection)returnValue;
            }
        }
  
        public void SaveAs(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public bool Saved 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Saved", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public vbext_ProjectType Type 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (vbext_ProjectType)returnValue;
            }
        }

        #endregion
    }
}
