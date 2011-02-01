using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Styles
{
    public class XlCellFormat : XlNonCreatable 
    {  
        #region Construction

        internal XlCellFormat(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{

		}

        #endregion       

        #region Methods

        public void Clear()
        {
            InstanceType.InvokeMember("Clear", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region Variant Properties

        public object AddIndent 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AddIndent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AddIndent", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object FormulaHidden 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FormulaHidden", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FormulaHidden", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object HorizontalAlignment 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HorizontalAlignment", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object IndentLevel 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("IndentLevel", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("IndentLevel", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object Locked 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Locked", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Locked", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object MergeCells 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MergeCells", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("MergeCells", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object NumberFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("NumberFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("NumberFormat", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object NumberFormatLocal 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("NumberFormatLocal", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("NumberFormatLocal", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object Orientation 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Orientation", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Orientation", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
 
        public object ShrinkToFit 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShrinkToFit", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ShrinkToFit", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object VerticalAlignment 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VerticalAlignment", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object WrapText 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WrapText", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("WrapText", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region COMReference Properties

        public XlApplication Application
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlApplication newClass = new XlApplication(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlBorders Borders 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Borders", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlBorders newClass = new XlBorders(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlFont Font 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Font", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlFont newClass = new XlFont(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlInterior Interior 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Interior", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlInterior newClass = new XlInterior(this, returnValue);
                if (null == returnValue) return null;
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public XlCreator Creator 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        #endregion
    }
}
