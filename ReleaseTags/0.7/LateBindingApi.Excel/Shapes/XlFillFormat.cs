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
    public class XlFillFormat : XlNonCreatable
    {  
        #region Construction

        internal XlFillFormat(IXlObject parentReference, object comReference) : base(parentReference, comReference)
		{
		    
		}

        #endregion

        #region Methods

        public void OneColorGradient(MsoGradientStyle style, long variant, Single degree)
        {
            object[] paramArray = new object[3];
            paramArray[0] = style;
            paramArray[1] = variant;
            paramArray[2] = degree;
            InstanceType.InvokeMember("OneColorGradient", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Patterned(MsoPatternType pattern)
        {
            object[] paramArray = new object[1];
            paramArray[0] = pattern;
            InstanceType.InvokeMember("Patterned", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PresetGradient(MsoGradientStyle style, long variant, MsoPresetGradientType presetGradientType)
        {
            object[] paramArray = new object[3];
            paramArray[0] = style;
            paramArray[1] = variant;
            paramArray[2] = presetGradientType;
            InstanceType.InvokeMember("PresetGradient", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PresetTextured(MsoPresetTexture presetTexture)
        {
            object[] paramArray = new object[1];
            paramArray[0] = presetTexture;
            InstanceType.InvokeMember("PresetTextured", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Solid()
        {
            InstanceType.InvokeMember("Solid", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }
        public void UserPicture(string pictureFile)
        {
            object[] paramArray = new object[1];
            paramArray[0] = pictureFile;
            InstanceType.InvokeMember("UserPicture", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void UserTextured(string textureFile)
        {
            object[] paramArray = new object[1];
            paramArray[0] = textureFile;
            InstanceType.InvokeMember("UserTextured", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlColorFormat BackColor
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BackColor", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlColorFormat newClass = new XlColorFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("BackColor", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlColorFormat ForeColor
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ForeColor", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlColorFormat newClass = new XlColorFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ForeColor", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region Scalar Properties

        public MsoGradientColorType GradientColorType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GradientColorType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoGradientColorType)returnValue;
            }
        }

        public Single GradientDegree
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GradientDegree", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
        }

        public MsoGradientStyle GradientStyle
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GradientStyle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoGradientStyle)returnValue;
            }
        }

        public int GradientVariant
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GradientVariant", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public MsoPatternType Pattern
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Pattern", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPatternType)returnValue;
            }
        }

        public MsoPresetGradientType PresetGradientType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PresetGradientType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPresetGradientType)returnValue;
            }
        }

        public MsoPresetTexture PresetTexture
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PresetTexture", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoPresetTexture)returnValue;
            }
        }
        public string TextureName
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TextureName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public MsoTextureType TextureType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TextureType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTextureType)returnValue;
            }
        }

        public Single Transparency
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Transparency", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Transparency", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public void TwoColorGradient(MsoGradientStyle style, long variant)
        {
            object[] paramArray = new object[1];
            paramArray[0] = style;
            paramArray[1] = variant;
            InstanceType.InvokeMember("TwoColorGradient", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public MsoFillType Type
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type ", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoFillType)returnValue;
            }
        }

        public MsoTriState Visible
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
