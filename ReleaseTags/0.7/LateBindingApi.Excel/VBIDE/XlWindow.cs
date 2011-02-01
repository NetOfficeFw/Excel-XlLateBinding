using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlWindow : XlNonCreatable
    {
        #region Construction

        internal XlWindow(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion

        #region COMReference Properties

        public XlWindows Collection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Collection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWindows newClass = new XlWindows(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlLinkedWindows LinkedWindows
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LinkedWindows", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlLinkedWindows newClass = new XlLinkedWindows(this, returnValue);
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

        public XlWindow LinkedWindowFrame
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LinkedWindowFrame", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWindow newClass = new XlWindow(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public string Caption 
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("Caption", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public int Height 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Height", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Left 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Left", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Left", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Top 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Top", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public vbext_WindowType Type  
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (vbext_WindowType)returnValue;
            }
        }
 
        public int Visible 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Width 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Width", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public vbext_WindowState WindowState 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WindowState", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (vbext_WindowState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("WindowState", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region Methods

        public void Close()
        {
            InstanceType.InvokeMember("Close", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetFocus()
        {
            InstanceType.InvokeMember("SetFocus", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion
    }
}
