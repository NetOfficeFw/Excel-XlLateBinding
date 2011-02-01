using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Text;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Validation
{
    /// <summary>
    /// Represents Validation in Excel
    /// </summary>
    public class XlValidation : XlNonCreatable
    {
        #region Construction

        internal XlValidation(IXlObject parentReference, object comReference):base(parentReference,comReference)
		{
		}

        #endregion

        #region Scalar Properties

        public bool IgnoreBlank
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("IgnoreBlank", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("IgnoreBlank", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool InCellDropdown
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("InCellDropdown", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("InCellDropdown", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowError
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("InCellDropdown", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("InCellDropdown", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region Methods

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);    
        }

        public void Add(XLXlDVType type, XLXlDVAlertStyle alertStyle, XLFormatConditionOperator Operator, string formular1)
        { 
            object[] parameter = new object[5];
            parameter[0] = type;
            parameter[1] = alertStyle;
            parameter[2] = Operator;
            parameter[3] = formular1;
            parameter[4] = Missing.Value;

            InstanceType.InvokeMember("Add", BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding,
            null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion
    }
}
