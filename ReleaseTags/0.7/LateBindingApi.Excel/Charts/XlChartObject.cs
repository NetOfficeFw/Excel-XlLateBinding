using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Charts
{

    public class XlChartObject : XlNonCreatable
    {
        #region Construction

        internal XlChartObject(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{			 
		}

        #endregion  

        #region Methods

        public void Activate()
        {
            InstanceType.InvokeMember("Activate", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }
        public void BringToFront()
        {
            InstanceType.InvokeMember("BringToFront", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Copy()
        {
            InstanceType.InvokeMember("Copy", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void CopyPicture()
        {
            InstanceType.InvokeMember("CopyPicture", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Cut()
        {
            InstanceType.InvokeMember("Cut", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }
        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SendToBack()
        {
            InstanceType.InvokeMember("SendToBack", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlChartObject Duplicate()
        {
            object returnValue  = InstanceType.InvokeMember("Duplicate", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartObject newClass = new XlChartObject(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
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

        public XlRange BottomRightCell
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BottomRightCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlChart Chart
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Chart", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChart newClass = new XlChart(this, returnValue);
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

        public XlShapeRange ShapeRange
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShapeRange", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShapeRange newClass = new XlShapeRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange TopLeftCell
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TopLeftCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties
 
        public bool Enabled
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Enabled", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Enabled", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Double Height
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Height", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Double)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Index
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Index", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
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

        public bool Locked
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Locked", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Locked", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Name", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool PrintObject
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PrintObject", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("PrintObject", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectChartObject
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectChartObject", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ProtectChartObject", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
        
        public bool RoundedCorners
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RoundedCorners", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("RoundedCorners", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        public bool Visible
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Double Width
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Width", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Double)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }   
        
        #endregion
    }

}
