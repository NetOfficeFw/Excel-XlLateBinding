using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Office;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlVBE : XlNonCreatable
    {
        #region Construction

        internal XlVBE(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion

        #region COMReference Properties
        public XlCodePane ActiveCodePane
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveCodePane", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCodePane newClass = new XlCodePane(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBProject ActiveVBProject
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveVBProject", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBProject newClass = new XlVBProject(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWindow ActiveWindow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveWindow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWindow newClass = new XlWindow(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlAddins Addins
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Addins", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlAddins newClass = new XlAddins(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCodePanes CodePanes
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CodePanes", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCodePanes newClass = new XlCodePanes(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCommandBars CommandBars
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CommandBars", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBars newClass = new XlCommandBars(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlEvents Events
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Events", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlEvents newClass = new XlEvents(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWindow MainWindow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MainWindow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWindow newClass = new XlWindow(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBComponent SelectedVBComponent
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SelectedVBComponent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBComponent newClass = new XlVBComponent(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBProjects VBProjects
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBProjects", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBProjects newClass = new XlVBProjects(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWindows Windows
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Windows", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWindows newClass = new XlWindows(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public string Version 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Version", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        #endregion
    }
}
