using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel
{
    public class XlProtection : XlNonCreatable
    {
        #region Construction

        internal XlProtection(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion
        
        #region COMReference Properties

        public XlAllowEditRanges AllowEditRanges 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowEditRanges", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);               
                if (null == returnValue) return null;
                XlAllowEditRanges newClass = new XlAllowEditRanges(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        
        #endregion

        #region Scalar Properties

        public bool AllowDeletingColumns 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowDeletingColumns", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowDeletingRows 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowDeletingRows", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowFiltering 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowFiltering", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }


        public bool AllowFormattingCells 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowFormattingCells", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowFormattingColumns 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowFormattingColumns", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowFormattingRows 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowFormattingRows", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowInsertingColumns 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowInsertingColumns", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowInsertingHyperlinks 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowInsertingHyperlinks", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowInsertingRows 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowInsertingRows", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowSorting 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowSorting", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool AllowUsingPivotTables 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AllowUsingPivotTables", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        #endregion
    }
}
