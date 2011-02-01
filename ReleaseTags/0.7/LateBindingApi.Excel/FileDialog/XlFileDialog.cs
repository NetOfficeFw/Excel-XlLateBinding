using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.FileDialog
{

    public class XlFileDialog : XlNonCreatable 
    {
        #region Construction

        internal XlFileDialog(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{		 
		}

        #endregion

        #region Methods

        public void Execute()
        {
            InstanceType.InvokeMember("Execute", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public int Show()
        {
            object returnValue  = InstanceType.InvokeMember("Show", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (int)returnValue;
        }

        #endregion

        #region COMReference Properties

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

        public XlFileDialogFilters Filters
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Filters", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlFileDialogFilters newClass = new XlFileDialogFilters(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlFileDialogSelectedItems SelectedItems 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SelectedItems", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlFileDialogSelectedItems newClass = new XlFileDialogSelectedItems(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public bool AllowMultiSelect
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AllowMultiSelect", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AllowMultiSelect", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string ButtonName
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ButtonName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ButtonName", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Creator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public MsoFileDialogType DialogType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DialogType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoFileDialogType)returnValue;
            }
        }

        public int FilterIndex
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FilterIndex", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FilterIndex", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string InitialFileName 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("InitialFileName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("InitialFileName", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoFileDialogView InitialView 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("InitialView", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoFileDialogView)returnValue;
            }
        }

        public string Item
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string Title 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Title", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Title", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }

}
