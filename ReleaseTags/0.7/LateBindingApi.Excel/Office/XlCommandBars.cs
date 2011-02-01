using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Office
{
    public class XlCommandBars : XlNonCreatable, IXlEventBinding, System.Collections.IEnumerable
    {   
        #region Fields

        XlCommandBarsEvents _eventBridge;

        #endregion

        #region Construction

        internal XlCommandBars(IXlObject parentReference, object comReference) : base(parentReference, comReference)
        {
        }

        #endregion 

        #region Methods

        public XlCommandBar Add(string name, MsoBarPosition position, object menuBar, object temporary)
        {
            object[] paramArray = new object[4];
            paramArray[0] = name;
            paramArray[1] = position;
            paramArray[2] = menuBar;
            paramArray[3] = temporary;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCommandBar newClass = new XlCommandBar(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }
        
        public void ReleaseFocus()
        {
            InstanceType.InvokeMember("ReleaseFocus", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlCommandBarControl FindControl(object type, object id, object tag, object visible)
        {
            object[] paramArray = new object[4];
            paramArray[0] = type;
            paramArray[1] = id;
            paramArray[2] = tag;
            paramArray[3] = visible;
            object returnValue  = InstanceType.InvokeMember("FindControl", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCommandBarControl newClass = new XlCommandBarControl(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlCommandBarControls FindControls(object type, object id, object tag, object visible)
        {
            object[] paramArray = new object[4];
            paramArray[0] = type;
            paramArray[1] = id;
            paramArray[2] = tag;
            paramArray[3] = visible;
            object returnValue  = InstanceType.InvokeMember("FindControl", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCommandBarControls newClass = new XlCommandBarControls(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        #endregion

        #region COMReference Properties

        /// <summary>
        /// returns an XlCommandBar by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlCommandBar this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBar newClass = new XlCommandBar(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        /// <summary>
        /// returns an XlCommandBar by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlCommandBar this[string index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBar newClass = new XlCommandBar(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCommandBarControl ActionControl
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActionControl", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBarControl newClass = new XlCommandBarControl(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCommandBar ActiveMenuBar
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveMenuBar", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBar newClass = new XlCommandBar(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

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

        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public bool AdaptiveMenus
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AdaptiveMenus", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AdaptiveMenus", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
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

        public bool DisableAskAQuestionDropDown 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisableAskAQuestionDropdown", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisableAskAQuestionDropdown", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisableCustomize 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisableCustomize", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisableCustomize", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayFonts 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayFonts", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayFonts", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayKeysInTooltips 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayKeysInTooltips", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayKeysInTooltips", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayTooltips 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayTooltips", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayTooltips", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool LargeButtons 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LargeButtons", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("LargeButtons", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoMenuAnimation MenuAnimationStyle 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MenuAnimationStyle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoMenuAnimation)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("MenuAnimationStyle", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
 
        #region Foreach

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            int iCount = Count;
            XlCommandBar[] res_addins = new XlCommandBar[iCount];

            for (int i = 1; i <= iCount; i++)
                res_addins[i - 1] = this[i];

            for (int i = 0; i < res_addins.Length; i++)
            {
                yield return res_addins[i];
            }

        }

        #endregion
     
        #region Events

        public event CommandBarsEvents_OnUpdateEventHandler OnUpdate;
        internal void RaiseOnUpdateEvent()
        {
            if (null == OnUpdate)
            {       
                return;
            }

            OnUpdate();
        }

        #endregion

        #region IXlEventBinding Members

        public void SetupEventBinding()
        {
            _eventBridge = new XlCommandBarsEvents();
            _eventBridge.SetupEventBinding(this);
        }

        public void RemoveEventBinding()
        {
            _eventBridge.RemoveEventBinding();
        }

        #endregion
    }
}
