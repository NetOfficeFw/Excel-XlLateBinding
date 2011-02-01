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
    public class XlDiagramNodeChildren : XlNonCreatable
    {
        #region Construction

        internal XlDiagramNodeChildren(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion

        #region Methods

        public XlDiagramNode AddNode(int index, MsoDiagramNodeType nodeType)
        {
            object[] paramArray = new object[2];
            paramArray[0] = index;
            paramArray[1] = nodeType;
            object returnValue  = InstanceType.InvokeMember("AddNode", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void SelectAll()
        {
            InstanceType.InvokeMember("SelectAll", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlDiagramNode FirstChild
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FirstChild", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDiagramNode Item(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlDiagramNode LastChild
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LastChild", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        #endregion

        #region Scalar Properties

        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion
    }
}
