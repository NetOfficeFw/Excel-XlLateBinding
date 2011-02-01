using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Charts
{
    public class XlOLEObjects : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlOLEObjects(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion

        #region Methods

        public XlOLEObject Add(string classType, string fileName, double left,double top,double width,double height)
        {
            object[] parameter = new object[11];
            parameter[0] = classType;
            parameter[1] = fileName;
            parameter[2] = Missing.Value;
            parameter[3] = Missing.Value;
            parameter[4] = Missing.Value;
            parameter[5] = Missing.Value;
            parameter[6] = Missing.Value;
            parameter[7] = left;
            parameter[8] = top;
            parameter[9] = width;
            parameter[10] = height;
            object returnValue = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlOLEObject newClass = new XlOLEObject(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

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

        /// <summary>
        /// returns an XlOLEObject by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlOLEObject this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlOLEObject newClass = new XlOLEObject(this, returnValue);
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

        public XlCreator Creator
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
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
        
        #region Foreach

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            int iCount = Count;
            XlOLEObject[] res_addins = new XlOLEObject[iCount];

            for (int i = 1; i <= iCount; i++)
                res_addins[i - 1] = this[i];

            for (int i = 0; i < res_addins.Length; i++)
            {
                yield return res_addins[i];
            }

        }

        #endregion
    }
}
