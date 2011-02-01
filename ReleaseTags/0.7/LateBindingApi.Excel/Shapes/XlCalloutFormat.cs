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
    public class XlCalloutFormat : XlNonCreatable
    {  
        #region Construction

        internal XlCalloutFormat(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion
 
        #region Methods

        public void AutomaticLength()
        {
            InstanceType.InvokeMember("AutomaticLength", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void CustomDrop(Single drop)
        {
            object[] paramArray = new object[1];
            paramArray[0] = drop;
            InstanceType.InvokeMember("CustomDrop", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void CustomLength(Single length)
        {
            object[] paramArray = new object[1];
            paramArray[0] = length;
            InstanceType.InvokeMember("CustomLength", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PresetDrop(MsoCalloutDropType dropType)
        {
            object[] paramArray = new object[1];
            paramArray[0] = dropType;
            InstanceType.InvokeMember("PresetDrop", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion
        
        #region Scalar Properties

        public MsoTriState Accent
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Accent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set 
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Accent", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoCalloutAngleType Angle
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Angle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoCalloutAngleType)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Angle", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTriState AutoAttach
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoAttach", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoAttach", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTriState AutoLength
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoLength", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoLength", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTriState Border 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Border", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Border", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single Drop
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Drop", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
        }

        public MsoCalloutDropType DropType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DropType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoCalloutDropType)returnValue;
            }
        }

        public Single Gap
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Gap", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
        }

        public Single Length
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Length", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
        }

        public MsoCalloutType Type
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoCalloutType)returnValue;
            }
        }

        #endregion
    }
}
