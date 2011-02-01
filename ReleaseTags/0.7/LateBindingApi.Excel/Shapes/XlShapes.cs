using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Shapes
{
    /// <summary>
    ///  Represents the Shapes Collection from Excel.Worksheet
    /// </summary>
    public class XlShapes : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlShapes(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion

        #region Methods

        public XlShape AddCallout(MsoCalloutType type, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[5];
            parameters[0] = type;
            parameters[1] = left;
            parameters[2] = top;
            parameters[3] = width;
            parameters[4] = height;

            object returnValue = InstanceType.InvokeMember("AddCallout", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddConnector(MsoConnectorType type, Single beginX, Single beginY, Single endX, Single endY)
        {
            object[] parameters = new object[5];
            parameters[0] = type;
            parameters[1] = beginX;
            parameters[2] = beginY;
            parameters[3] = endX;
            parameters[4] = endY;
            object returnValue = InstanceType.InvokeMember("AddConnector", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddCurve(Array safeArrayOfPoints)
        {
            object[] parameters = new object[1];
            parameters[0] = safeArrayOfPoints;
            object returnValue = InstanceType.InvokeMember("AddCurve", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddDiagram(MsoDiagramType type, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[5];
            parameters[0] = type;
            parameters[1] = left;
            parameters[2] = top;
            parameters[3] = width;
            parameters[4] = height;
            object returnValue = InstanceType.InvokeMember("AddDiagram", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddFormControl(XlFormControl type, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[5];
            parameters[0] = type;
            parameters[1] = left;
            parameters[2] = top;
            parameters[3] = width;
            parameters[4] = height;
            object returnValue = InstanceType.InvokeMember("AddFormControl", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddLabel(MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[5];
            parameters[0] = orientation;
            parameters[1] = left;
            parameters[2] = top;
            parameters[3] = width;
            parameters[4] = height;
            object returnValue = InstanceType.InvokeMember("AddLabel", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddLine(Single beginX, Single beginY, Single endX, Single endY)
        {
            object[] parameters = new object[4];
            parameters[0] = beginX;
            parameters[1] = beginY;
            parameters[2] = endX;
            parameters[3] = endY;
            object returnValue = InstanceType.InvokeMember("AddLine", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddOLEObject(string classType, string filename, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[11];
            parameters[0] = classType;
            parameters[1] = filename;
            parameters[2] = Missing.Value;
            parameters[3] = Missing.Value;
            parameters[4] = Missing.Value;
            parameters[5] = Missing.Value;
            parameters[6] = Missing.Value;
            parameters[7] = left;
            parameters[8] = top;
            parameters[9] = width;
            parameters[10] = height;
            object returnValue = InstanceType.InvokeMember("AddOLEObject", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddPicture(string filename, MsoTriState linkToFile, MsoTriState saveWithDocument, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[7];
            parameters[0] = filename;
            parameters[1] = linkToFile;
            parameters[2] = saveWithDocument;
            parameters[3] = left;
            parameters[4] = top;
            parameters[5] = width;
            parameters[6] = height;
            object returnValue = InstanceType.InvokeMember("AddPicture", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddPolyline(Array safeArrayOfPoints)
        {
            object[] parameters = new object[1];
            parameters[0] = safeArrayOfPoints;
            object returnValue = InstanceType.InvokeMember("AddPolyline", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddShape(MsoAutoShapeType type, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[5];
            parameters[0] = type;
            parameters[1] = left;
            parameters[2] = top;
            parameters[3] = width;
            parameters[4] = height;
            object returnValue = InstanceType.InvokeMember("AddShape", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddTextbox(MsoTextOrientation orientation, Single left, Single top, Single width, Single height)
        {
            object[] parameters = new object[5];
            parameters[0] = orientation;
            parameters[1] = left;
            parameters[2] = top;
            parameters[3] = width;
            parameters[4] = height;
            object returnValue = InstanceType.InvokeMember("AddTextbox", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShape AddTextEffect(MsoPresetTextEffect presetTextEffect, string text, string fontName, Single fontSize, MsoTriState fontBold, MsoTriState fontItalic, Single left, Single top)
        {
            object[] parameters = new object[8];
            parameters[0] = presetTextEffect;
            parameters[1] = text;
            parameters[2] = fontName;
            parameters[3] = fontSize;
            parameters[4] = fontBold;
            parameters[5] = fontItalic;
            parameters[6] = left;
            parameters[7] = top;

            object returnValue = InstanceType.InvokeMember("AddTextEffect", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlFreeformBuilder BuildFreeform(MsoEditingType editingType, Single x1, Single y1)
        {
            object[] parameters = new object[3];
            parameters[0] = editingType;
            parameters[1] = x1;
            parameters[2] = y1;
            object returnValue = InstanceType.InvokeMember("BuildFreeform", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, parameters, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlFreeformBuilder newClass = new XlFreeformBuilder(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void SelectAll()
        {
            InstanceType.InvokeMember("SelectAll", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region Scalar Properties

        /// <summary>
        /// Count of Shapes
        /// </summary>
        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion

        #region ComReference Properties

        /// <summary>
        /// returns a Shape by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlShape this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShape newClass = new XlShape(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region ForEach

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {

            /*
             *
             * yield cant have a try catch block
             * 
            */

            int iCount = Count;
            XlShape[] res_shapes = new XlShape[iCount];

            for (int i = 1; i <= iCount; i++)
                res_shapes[i - 1] = this[i];

            for (int i = 0; i < res_shapes.Length; i++)
            {
                yield return res_shapes[i];
            }


        }

        #endregion
    }
}
