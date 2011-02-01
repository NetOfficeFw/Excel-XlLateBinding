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
    public class XlDiagramNode : XlNonCreatable
    {   
        #region Construction

        internal XlDiagramNode(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion

        #region Methods

        public XlDiagramNode AddNode(MsoRelativeNodePosition mso, MsoDiagramNodeType nodeType)
        {
            object[] paramArray = new object[2];
            paramArray[0] = mso;
            paramArray[1] = nodeType;
            object returnValue  = InstanceType.InvokeMember("AddNode", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null; 
            XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlDiagramNode CloneNode(bool copyChildren, XlDiagramNode targetNode, MsoRelativeNodePosition pos)
        {
            object[] paramArray = new object[3];
            paramArray[0] = copyChildren;
            paramArray[1] = targetNode.COMReference;
            paramArray[2] = pos;
            object returnValue  = InstanceType.InvokeMember("CloneNode", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Delete()
        {
            InstanceType.InvokeMember("CloneNode", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void MoveNode(XlDiagramNode targetNode, MsoRelativeNodePosition pos)
        {
            object[] paramArray = new object[2];
            paramArray[0] = targetNode.COMReference;
            paramArray[1] = pos;
            InstanceType.InvokeMember("MoveNode", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlDiagramNode NextNode()
        {
            object returnValue  = InstanceType.InvokeMember("NextNode", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlDiagramNode PrevNode()
        {
            object returnValue  = InstanceType.InvokeMember("PrevNode", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
           if (null == returnValue) return null;
            XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void ReplaceNode(XlDiagramNode targetNode)
        {
            object[] paramArray = new object[1];
            paramArray[0] = targetNode.COMReference;
            InstanceType.InvokeMember("pTargetNode", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SwapNode(XlDiagramNode targetNode, bool swapChildren)
        {
            object[] paramArray = new object[2];
            paramArray[0] = targetNode.COMReference;
            paramArray[1] = swapChildren;
            InstanceType.InvokeMember("SwapNode", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void TransferChildren(XlDiagramNode receivingNode)
        {
            object[] paramArray = new object[1];
            paramArray[0] = receivingNode.COMReference;
            InstanceType.InvokeMember("TransferChildren", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlDiagramNodeChildren Children
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Children", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDiagramNodeChildren newClass = new XlDiagramNodeChildren(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDiagramNode Root
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Root", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlShape Shape
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Shape", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShape newClass = new XlShape(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlShape TextShape
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TextShape", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShape newClass = new XlShape(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public MsoOrgChartLayoutType Layout
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Layout", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoOrgChartLayoutType)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Layout", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }
}
