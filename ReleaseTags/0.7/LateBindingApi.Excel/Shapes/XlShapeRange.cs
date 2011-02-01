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
    public class XlShapeRange : XlNonCreatable
    {   
        #region Construction

        internal XlShapeRange(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion

        #region Methods

        public void Align(MsoAlignCmd alignCmd, MsoTriState relativeTo)
        {
            object[] paramArray = new object[2];
            paramArray[0] = alignCmd;
            paramArray[1] = relativeTo;
            InstanceType.InvokeMember("Align", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Apply()
        {
            InstanceType.InvokeMember("Apply", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Distribute(MsoDistributeCmd distributeCmd, MsoTriState relativeTo)
        {
            object[] paramArray = new object[2];
            paramArray[0] = distributeCmd;
            paramArray[1] = relativeTo;
            InstanceType.InvokeMember("Distribute", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlShapeRange Duplicate()
        {
            object returnValue  = InstanceType.InvokeMember("Duplicate", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShapeRange newClass = new XlShapeRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Flip(MsoFlipCmd flipCmd)
        {
            object[] paramArray = new object[1];
            paramArray[0] = flipCmd;
            InstanceType.InvokeMember("Flip", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlShape Group()
        {
            object returnValue  = InstanceType.InvokeMember("Group", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void IncrementLeft(Single increment)
        {
            object[] paramArray = new object[1];
            paramArray[0] = increment;
            InstanceType.InvokeMember("IncrementLeft", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void IncrementRotation(Single increment)
        {
            object[] paramArray = new object[1];
            paramArray[0] = increment;
            InstanceType.InvokeMember("IncrementRotation", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void IncrementTop(Single increment)
        {
            object[] paramArray = new object[1];
            paramArray[0] = increment;
            InstanceType.InvokeMember("IncrementTop", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlShape Item(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void PickUp()
        {
            InstanceType.InvokeMember("PickUp", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlShape Regroup()
        {
            object returnValue  = InstanceType.InvokeMember("Regroup", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void RerouteConnections()
        {
            InstanceType.InvokeMember("RerouteConnections", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ScaleHeight(Single factor, MsoTriState relativeToOriginalSize)
        {
            object[] paramArray = new object[2];
            paramArray[0] = factor;
            paramArray[1] = relativeToOriginalSize;
            InstanceType.InvokeMember("ScaleHeight", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ScaleWidth(Single factor, MsoTriState relativeToOriginalSize)
        {
            object[] paramArray = new object[2];
            paramArray[0] = factor;
            paramArray[1] = relativeToOriginalSize;
            InstanceType.InvokeMember("ScaleWidth", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetShapesDefaultProperties()
        {
            InstanceType.InvokeMember("SetShapesDefaultProperties", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlShapeRange Ungroup()
        {
            object returnValue  = InstanceType.InvokeMember("Ungroup", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShapeRange newClass = new XlShapeRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        #endregion

        #region COMReference Properties

        public XlAdjustments Adjustments
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Adjustments", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlAdjustments newClass = new XlAdjustments(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlCalloutFormat Callout
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Callout", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCalloutFormat newClass = new XlCalloutFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlConnectorFormat ConnectorFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ConnectorFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlConnectorFormat newClass = new XlConnectorFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDiagram Diagram
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Diagram", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDiagram newClass = new XlDiagram(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDiagramNode DiagramNode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DiagramNode", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlFillFormat Fill
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Fill", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlFillFormat newClass = new XlFillFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlGroupShapes GroupItems
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GroupItems", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlGroupShapes newClass = new XlGroupShapes(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlLineFormat Line
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Line", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlLineFormat newClass = new XlLineFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlShapeNodes Nodes
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Nodes", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShapeNodes newClass = new XlShapeNodes(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlShape ParentGroup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ParentGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShape newClass = new XlShape(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }


        public XlPictureFormat PictureFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PictureFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPictureFormat newClass = new XlPictureFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlThreeDFormat ThreeD
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ThreeD", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlThreeDFormat newClass = new XlThreeDFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }


        public XlShadowFormat Shadow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Shadow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShadowFormat newClass = new XlShadowFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlTextEffectFormat TextEffect
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TextEffect", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlTextEffectFormat newClass = new XlTextEffectFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlTextFrame TextFrame
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TextFrame", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlTextFrame newClass = new XlTextFrame(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        

        #endregion

        #region Scalar Properties

        public string AlternativeText
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AlternativeText", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AlternativeText", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoAutoShapeType AutoShapeType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoShapeType", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoAutoShapeType)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoShapeType", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoBlackWhiteMode BlackWhiteMode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BlackWhiteMode", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoBlackWhiteMode)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("BlackWhiteMode", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTriState Child 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Child", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public int ConnectionSiteCount
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ConnectionSiteCount", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public MsoTriState Connector
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Connector", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public MsoTriState HasDiagram
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasDiagram", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public MsoTriState HasDiagramNode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasDiagramNode", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }
        
        public Single Height
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("Height", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTriState HorizontalFlip
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HorizontalFlip", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public int ID
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ID", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public Single Left
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("Left", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Left", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoTriState LockAspectRatio
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("LockAspectRatio", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("LockAspectRatio", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        public Single Rotation
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("Rotation", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Rotation", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single Top
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("Top", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoShapeType Type
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoShapeType)returnValue;
            }
        }

        public MsoTriState VerticalFlip
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("VerticalFlip", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public MsoTriState Visible
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("Visible ", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public Single Width
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("Width", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int ZOrderPosition
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("ZOrderPosition", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }
        
        #endregion
    }
}
