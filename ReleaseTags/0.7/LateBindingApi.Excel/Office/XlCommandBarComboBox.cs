using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;
using System.Security;
using System.Security.Permissions;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Stdole;

namespace LateBindingApi.Excel.Office
{
    public class XlCommandBarComboBox : XlCommandBarControl, IXlEventBinding 
    {
        #region Fields

        XlCommandBarComboBoxEvents _eventBridge;

        #endregion

        #region Construction

        internal XlCommandBarComboBox(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		
		}

        #endregion 

        #region Methods

        public void Clear()
        {
            InstanceType.InvokeMember("Clear", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void AddItem(string text)
        {
            object[] paramArray = new object[2];
            paramArray[0] = text;
            paramArray[1] = Missing.Value;
            InstanceType.InvokeMember("AddItem", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void AddItem(string text, int index)
        {
            object[] paramArray = new object[2];
            paramArray[0] = text;
            paramArray[1] = index;
            InstanceType.InvokeMember("AddItem", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RemoveItem(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            InstanceType.InvokeMember("RemoveItem", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);

        }

        #endregion

        #region Scalar Properties

        public int ListCount 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ListCount", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public string List(int index)
        {
            object[] parameter = new object[1];
            parameter[0] = index;
            object returnValue  = InstanceType.InvokeMember("List", BindingFlags.GetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            return (string)returnValue;
        }

        public int ListHeaderCount 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ListHeaderCount", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ListHeaderCount", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region Events

        public event CommandBarComboBoxEvents_ChangeEventHandler Change;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseChangeEvent(object Ctrl)
        {
            if (null == Change)
            {
                Marshal.ReleaseComObject(Ctrl);
                return;
            }

            XlCommandBarComboBox box = new XlCommandBarComboBox(this, Ctrl);
            ListChildReferences.Add(box);

            Change(box);
        }

        #endregion

        #region IXlEventBinding Members

        public void SetupEventBinding()
        {
            _eventBridge = new XlCommandBarComboBoxEvents();
            _eventBridge.SetupEventBinding(this);
        }

        public void RemoveEventBinding()
        {
            _eventBridge.RemoveEventBinding();
        }

        #endregion
    }
}
