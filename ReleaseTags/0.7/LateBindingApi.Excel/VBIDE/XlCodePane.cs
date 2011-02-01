using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlCodePane : XlNonCreatable
    {
        #region Construction

        internal XlCodePane(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
		 
		}

        #endregion    

        #region COMReference Properties
        
        public XlCodeModule CodeModule
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CodeModule", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCodeModule newClass = new XlCodeModule(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCodePanes Collection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Collection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCodePanes newClass = new XlCodePanes(this, returnValue);
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

        public XlWindow Window
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Window", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWindow newClass = new XlWindow(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public vbext_CodePaneview CodePaneView 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CodePaneView", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (vbext_CodePaneview)returnValue;
            }
        }

        public int CountOfVisibleLines 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CountOfVisibleLines", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int TopLine 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TopLine", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("TopLine", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
  
        #endregion

        #region Methods

        public void GetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            object[] parameter = new object[4];
            parameter[0] = startLine;
            parameter[0] = startColumn;
            parameter[0] = endLine;
            parameter[0] = endColumn;
            InstanceType.InvokeMember("GetSelection", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetSelection(int startLine, int startColumn, int endLine, int endColumn)
        {
            object[] parameter = new object[4];
            parameter[0] = startLine;
            parameter[0] = startColumn;
            parameter[0] = endLine;
            parameter[0] = endColumn;
            InstanceType.InvokeMember("SetSelection", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Show()
        {
            InstanceType.InvokeMember("Show", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion
    }
}
