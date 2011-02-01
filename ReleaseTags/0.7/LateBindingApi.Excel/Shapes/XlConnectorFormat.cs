using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Shapes
{
    public class XlConnectorFormat : XlNonCreatable
    { 
        #region Construction

        internal XlConnectorFormat(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		    
        
        }

        #endregion

        #region Methods

        public void BeginConnect(XlShape connectedShape, int connectionSite)
        {
            object[] paramArray = new object[2];
            paramArray[0] = connectedShape.COMReference;
            paramArray[1] = connectionSite;
            InstanceType.InvokeMember("BeginConnect", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void BeginDisconnect()
        {
            InstanceType.InvokeMember("BeginConnect", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void EndConnect(XlShape connectedShape, int connectionSite)
        {
            object[] paramArray = new object[2];
            paramArray[0] = connectedShape.COMReference;
            paramArray[1] = connectionSite;
            InstanceType.InvokeMember("EndConnect", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlShape BeginConnectedShape
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BeginConnectedShape", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShape newClass = new XlShape(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public MsoTriState BeginConnected
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BeginConnected", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public int BeginConnectionSite
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BeginConnectionSite", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }
        public MsoTriState EndConnected
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EndConnected", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public XlShape EndConnectedShape
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EndConnectedShape", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlShape newClass = new XlShape(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public int EndConnectionSite
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EndConnectionSite", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public void EndDisconnect()
        {
            InstanceType.InvokeMember("EndDisconnect", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public MsoConnectorType Type
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoConnectorType)returnValue;
            }
        }

        #endregion
    }
}
