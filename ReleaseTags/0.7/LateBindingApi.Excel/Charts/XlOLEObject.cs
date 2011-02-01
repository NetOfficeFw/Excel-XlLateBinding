using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Styles; 
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Charts
{
    public class XlOLEObject : XlNonCreatable
    {
        #region Construction

        internal XlOLEObject(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		}

        #endregion

        #region Methods

        public void Activate()
        {
            InstanceType.InvokeMember("Activate", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void BringToFront()
        {
            InstanceType.InvokeMember("BringToFront", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);

        }

        public void Copy()
        {
            InstanceType.InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void CopyPicture()
        {
            object[] parameter = new object[2];
            parameter[0] = XlPictureAppearance.xlPrinter;
            parameter[1] = XlCopyPictureFormat.xlPicture;
            InstanceType.InvokeMember("CopyPicture", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }
        
        public void CopyPicture(XlPictureAppearance appearance, XlCopyPictureFormat format)
        {
            object[] parameter = new object[2];
            parameter[0] = appearance;
            parameter[1] = format;
            InstanceType.InvokeMember("CopyPicture", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Cut()
        {
            InstanceType.InvokeMember("Cut", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlOLEObject Duplicate()
        {
            object returnValue = InstanceType.InvokeMember("Duplicate", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);       
            if (null == returnValue) return null;
            XlOLEObject newClass = new XlOLEObject(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SendToBack()
        {
            InstanceType.InvokeMember("SendToBack", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Update()
        {
            InstanceType.InvokeMember("Update", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Verb()
        {
            object[] parameter = new object[1];
            parameter[0] = XlOLEVerb.xlVerbPrimary;
            InstanceType.InvokeMember("Verb", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Verb(XlOLEVerb verb)
        {
            object[] parameter = new object[1];
            parameter[0] = verb;
            InstanceType.InvokeMember("Verb", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlApplication Application
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlApplication newClass = new XlApplication(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlBorder Border
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Border", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue = InstanceType.InvokeMember("BottomRightCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlInterior Interior 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Interior", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlInterior newClass = new XlInterior(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlNonCreatable Parent
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Parent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlNonCreatable newClass = XlDynamicType.CreateDynamicType(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlShapeRange ShapeRange 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ShapeRange", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue = InstanceType.InvokeMember("TopLeftCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public bool AutoLoad 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AutoLoad", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AutoLoad", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AutoUpdate 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AutoUpdate", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AutoUpdate", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlCreator Creator 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public bool Enabled 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Enabled", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Enabled", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Height 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Height", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Index 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Index", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public double Left
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Left", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Left", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string LinkedCell 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("LinkedCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("LinkedCell", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string ListFillRange 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ListFillRange", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ListFillRange", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Locked 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Locked", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Locked", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Name 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Name", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Name", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object Object
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Object", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (object)returnValue;
            }
        }

        public bool PrintObject 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("PrintObject", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("PrintObject", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string ProgID 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("progID", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }           
        }

        public bool Shadow 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Shadow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Shadow", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string SourceName 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("SourceName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("SourceName", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Top 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Top", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Visible 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Width 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Width", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int ZOrder 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ZOrder", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion
    }
}
