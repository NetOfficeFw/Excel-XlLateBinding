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
  
    public class XlCommandBarButton : XlCommandBarControl, IXlEventBinding 
    {
        #region Fields

        XlCommandBarButtonEvents _eventBridge;

        #endregion

        #region Construction

        internal XlCommandBarButton(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		
		}

        #endregion 

        #region Methods

        public void CopyFace()
        {
            InstanceType.InvokeMember("CopyFace", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PasteFace()
        {
            InstanceType.InvokeMember("PasteFace", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlCommandBarControls Controls
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Controls", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBarControls newClass = new XlCommandBarControls(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        #endregion

        #region Scalar Properties

        public MsoButtonState State
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("State", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoButtonState)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("State", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoButtonStyle Style 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Style", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoButtonStyle)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Style", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int FaceId
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FaceId", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FaceId", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object Picture
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Picture", BindingFlags.GetProperty , null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                IPictureDisp newClass = new IPictureDisp(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;                
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Picture", BindingFlags.SetProperty | BindingFlags.OptionalParamBinding, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region Events
  
        public event CommandBarButtonEvents_ClickEventHandler Click;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseClickEvent(object Ctrl, ref bool CancelDefault)
        {
            if (null == Click)
            {
                Marshal.ReleaseComObject(Ctrl);
                return;
            }

            XlCommandBarButton btn = new XlCommandBarButton(this, Ctrl);
             ListChildReferences.Add(btn);
           
            Click(btn, ref CancelDefault);
        }

        #endregion

        #region IXlEventBinding Members

        public void SetupEventBinding()
        {
            _eventBridge = new XlCommandBarButtonEvents();
            _eventBridge.SetupEventBinding(this);
        }

        public void RemoveEventBinding()
        {
            _eventBridge.RemoveEventBinding();
        }

        #endregion
    }
}
