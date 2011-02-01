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
    public class XlThreeDFormat : XlNonCreatable
    {   
        #region Construction

        internal XlThreeDFormat(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion

        #region Methods

        public void IncrementRotationX(Single increment)
        {
            object[] paramArray = new object[1];
            paramArray[0] = increment;
            InstanceType.InvokeMember("IncrementRotationX", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void IncrementRotationY(Single increment)
        {
            object[] paramArray = new object[1];
            paramArray[0] = increment;
            InstanceType.InvokeMember("IncrementRotationY", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ResetRotation()
        {
            InstanceType.InvokeMember("ResetRotation", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetExtrusionDirection(MsoPresetExtrusionDirection presetExtrusionDirection)
        {
            object[] paramArray = new object[1];
            paramArray[0] = presetExtrusionDirection;
            InstanceType.InvokeMember("SetExtrusionDirection", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetThreeDFormat(MsoPresetThreeDFormat presetThreeDFormat)
        {
            object[] paramArray = new object[1];
            paramArray[0] = presetThreeDFormat;
            InstanceType.InvokeMember("SetThreeDFormat", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlColorFormat ExtrusionColor
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ExtrusionColor", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlColorFormat newClass = new XlColorFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public Single Depth
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Depth", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Depth", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
        public MsoExtrusionColorType ExtrusionColorType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ExtrusionColorType", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoExtrusionColorType)returnValue;
            }
        }

        public MsoTriState Perspective
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Perspective", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Perspective", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoPresetExtrusionDirection PresetExtrusionDirection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PresetExtrusionDirection", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPresetExtrusionDirection)returnValue;
            }
        }

        public MsoPresetLightingDirection PresetLightingDirection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PresetLightingDirection", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPresetLightingDirection)returnValue;
            }
        }

        public MsoPresetLightingSoftness PresetLightingSoftness
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PresetLightingSoftness", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPresetLightingSoftness)returnValue;
            }
        }

        public MsoPresetMaterial PresetMaterial
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PresetMaterial", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPresetMaterial)returnValue;
            }
        }

        public MsoPresetThreeDFormat PresetThreeDFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PresetThreeDFormat", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPresetThreeDFormat)returnValue;
            }
        }

        public Single RotationX
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RotationX", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("RotationX", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single RotationY
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RotationY", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("RotationY", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTriState Visible
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }
}
