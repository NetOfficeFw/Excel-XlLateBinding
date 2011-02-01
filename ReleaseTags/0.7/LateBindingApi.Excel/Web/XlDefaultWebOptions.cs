using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Web
{
    public class XlDefaultWebOptions : XlNonCreatable 
    {   
        #region Construction

        internal XlDefaultWebOptions(IXlObject parentReference, object comReference):base(parentReference,comReference)
		{		 
		}

        #endregion    

        #region COMReferences

        public XlApplication Application
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlApplication newClass = new XlApplication(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        #endregion
        
        #region Scalar Properties

        public bool AllowPNG
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AllowPNG", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AllowPNG", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AlwaysSaveInDefaultEncoding 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AlwaysSaveInDefaultEncoding", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AlwaysSaveInDefaultEncoding", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool CheckIfOfficeIsHTMLEditor 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CheckIfOfficeIsHTMLEditor", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("CheckIfOfficeIsHTMLEditor", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlCreator Creator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public bool DownloadComponents 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DownloadComponents", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("DownloadComponents", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoEncoding Encoding
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Encoding", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoEncoding)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Encoding", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string FolderSuffix
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FolderSuffix", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        #endregion
    }
}
