using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Shapes
{
    public class XlTextFrame : XlNonCreatable
    {
        #region Construction

        internal XlTextFrame(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion 

        #region Methods

        public XlCharacters Characters()
        {
            object returnValue  = InstanceType.InvokeMember("Characters", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCharacters newClass = new XlCharacters(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlCharacters Characters(int start, int length)
        {
            object[] paramArray = new object[2];
            paramArray[0] = start;
            paramArray[1] = length;
            object returnValue  = InstanceType.InvokeMember("Characters", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCharacters newClass = new XlCharacters(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        #endregion

        #region Scalar Properties

        public bool AutoMargins
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoMargins", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoMargins", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AutoSize
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoSize", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoSize", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlHAlign HorizontalAlignment
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HorizontalAlignment", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlHAlign)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlVAlign VerticalAlignment
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VerticalAlignment", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlVAlign)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single MarginBottom
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MarginBottom", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("MarginBottom", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single MarginLeft
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MarginLeft", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("MarginLeft", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single MarginRight
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MarginRight", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("MarginRight", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single MarginTop
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MarginTop", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("MarginTop", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTextOrientation Orientation
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Orientation", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTextOrientation)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Orientation", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int ReadingOrder
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ReadingOrder", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ReadingOrder", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion   
    }
}
