using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Shapes
{
    /// Represents a Shape in a Worksheet, Button, PictureBox Whatever...
    /// </summary>
    public class XlShape : XlNonCreatable
    {
        #region Construction

        internal XlShape(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion 
               
        #region Methods

        public void PickUp()
        {
            InstanceType.InvokeMember("IncrementLeft", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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

        public void Copy()
        {
            InstanceType.InvokeMember("Copy", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Cut()
        {
            InstanceType.InvokeMember("Cut", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlShape Duplicate()
        {
            object returnValue  = InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Apply()
        {
            InstanceType.InvokeMember("Apply", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }


        public void Flip(MsoFlipCmd flipCmd)
        {
            object[] paramArray = new object[1];
            paramArray[0] = flipCmd;
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RerouteConnections()
        {
            InstanceType.InvokeMember("RerouteConnections", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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

        public XlHyperlink Hyperlink
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Hyperlink ", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlHyperlink newClass = new XlHyperlink(this, returnValue);
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

        public XlLinkFormat LinkFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LinkFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlLinkFormat newClass = new XlLinkFormat(this, returnValue);
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

        public XlControlFormat ControlFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ControlFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlControlFormat newClass = new XlControlFormat(this, returnValue);
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
                object returnValue  = InstanceType.InvokeMember("XlFillFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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

        public XlTextFrame TextFrame
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TextFrame", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                XlTextFrame newClass = new XlTextFrame(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

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

        public XlCalloutFormat CalloutFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CalloutFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCalloutFormat newClass = new XlCalloutFormat(this, returnValue);
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

        public XlOLEFormat OLEFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("OLEFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlOLEFormat newClass = new XlOLEFormat(this, returnValue);
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

        public XlFormControl FormControlType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FormControlType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlFormControl)returnValue;
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

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }  
        }

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

        public MsoTriState Visible
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public float Left
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Left", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (float)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Left", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public float Top
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Top", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (float)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public float Height
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Height", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (float)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public float Width
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Width", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (float)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoAutoShapeType AutoShapeType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoShapeType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("BlackWhiteMode", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("Child", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public int ConnectionSiteCount
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ConnectionSiteCount", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public MsoTriState Connector
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Connector", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }
 
        public MsoTriState HasDiagram
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasDiagram", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public MsoTriState HasDiagramNode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasDiagramNode", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public MsoTriState HorizontalFlip
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HorizontalFlip", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public MsoTriState LockAspectRatio
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LockAspectRatio", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public bool Locked
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Locked", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Locked", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string OnAction
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("OnAction", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("OnAction", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlPlacement Placement
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Placement", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlPlacement)returnValue;
            }
        }
       
        public Single Rotation
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Rotation", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (Single)returnValue;
            }
        }

        public MsoShapeType Type
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoShapeType)returnValue;
            }
        }

        public MsoTriState VerticalFlip
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VerticalFlip", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoTriState)returnValue;
            }
        }

        public int ZOrderPosition
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ZOrderPosition", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion
    }
}
