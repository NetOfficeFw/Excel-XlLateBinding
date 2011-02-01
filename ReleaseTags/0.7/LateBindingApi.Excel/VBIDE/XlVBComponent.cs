using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlVBComponent : XlNonCreatable
    {    
        #region Construction

        internal XlVBComponent(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion

        #region Methods

        public void Activate()
        {
            InstanceType.InvokeMember("Activate", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlWindow DesignerWindow()
        {
            object returnValue  = InstanceType.InvokeMember("DesignerWindow", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWindow newClass = new XlWindow(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Export(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("Export", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        public XlVBComponents Collection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Collection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBComponents newClass = new XlVBComponents(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlProperties Properties
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Properties", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlProperties newClass = new XlProperties(this, returnValue);
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

        public string DesignerID 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DesignerID", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public bool HasOpenDesigner 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasOpenDesigner", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set 
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Name", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);

            }
        }

        public bool Saved 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Saved", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public vbext_ComponentType Type 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (vbext_ComponentType)returnValue;
            }
        }

        #endregion
    }
}
