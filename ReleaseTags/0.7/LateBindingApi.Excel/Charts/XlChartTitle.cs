using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Charts
{
    public class XlChartTitle : XlNonCreatable 
    {
        #region Construction

        internal XlChartTitle(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{

		}

        #endregion

        #region Methods

        public XlCharacters Characters()
        {
            object returnValue  = InstanceType.InvokeMember("Characters", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
            object returnValue  = InstanceType.InvokeMember("Characters", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCharacters newClass = new XlCharacters(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties
        public XlBorder Border
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Border", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlBorder newClass = new XlBorder(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartFillFormat Fill
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Fill", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartFillFormat newClass = new XlChartFillFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlFont Font
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Font", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlFont newClass = new XlFont(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlInterior Interior
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Interior", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlInterior newClass = new XlInterior(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public string Caption
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Caption", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Caption", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        public Double Left
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Left", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Double)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Left", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public int ReadingOrder
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ReadingOrder", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public bool Shadow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Shadow", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Shadow", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Text
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Text", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Text", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Double Top
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Top", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Double)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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
                InstanceType.InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }
}
