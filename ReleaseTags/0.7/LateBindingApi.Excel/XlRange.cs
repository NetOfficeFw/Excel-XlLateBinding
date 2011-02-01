using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Error; 
using LateBindingApi.Excel.Validation;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel
{
    /// <summary>
    /// Represents an ExcelRange
    /// </summary>   
    public class XlRange : XlNonCreatable 
    {       
        #region Construction
       
        internal XlRange(IXlObject parentReference, object comReference):base(parentReference,comReference)
		{			
		}

        #endregion

        #region Scalar Properties

        public bool AddIndent
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AddIndent", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AddIndent", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool MergeCells
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MergeCells", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("MergeCells", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Address
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Address", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue; 
            }
        }

        public string Text
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Text", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public XlHAlign HorizontalAlignment 
        {
             get
            {
                object returnValue  = InstanceType.InvokeMember("HorizontalAlignment", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlHAlign)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("HorizontalAlignment", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlVAlign VerticalAlignment
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VerticalAlignment", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlVAlign)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("VerticalAlignment", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string NumberFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("NumberFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {               
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("NumberFormat", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture); 
            }
        }

        public string NumberFormatLocal
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("NumberFormatLocal", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("NumberFormatLocal", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double RowHeight
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RowHeight", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("RowHeight", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double ColumnWidth
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ColumnWidth", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ColumnWidth", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Height
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Height", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {               
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Width
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Width", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public object Value
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Value", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Value", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Formula 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Formula", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Formula", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Top
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Top", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Left
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Left", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Left", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string FormulaLocal 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FormulaLocal", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FormulaLocal", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string FormulaR1C1 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FormulaR1C1", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FormulaR1C1", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string FormulaR1C1Local 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FormulaR1C1Local", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FormulaR1C1Local", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool WrapText
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WrapText", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("WrapText", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Hidden
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Hidden", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Hidden", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
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

        public int Row
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Row", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }      
        }

        public int Column
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Column", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion

        #region ComReference Properties

        public XlApplication Application
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlApplication newClass = new XlApplication(this, returnValue);
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

        public XlFont Font
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Font", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlFont newClass = new XlFont(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlBorders Borders
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Borders", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlBorders newClass = new XlBorders(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlValidation Validation
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Validation", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlValidation newClass = new XlValidation(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        public XlStyle Style
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Style", BindingFlags.GetProperty , null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlStyle newClass = new XlStyle(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value.COMReference;
                InstanceType.InvokeMember("Style", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string StyleAsString
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Style", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                string returnValueAsString = (string)returnValue.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, returnValue, null, XlLateBindingApiSettings.XlThreadCulture);
                Marshal.ReleaseComObject(returnValue);
                return returnValueAsString; 
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Style", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlRange Rows
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Rows", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Columns
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Columns", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Cells
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Cells", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange CurrentArray 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CurrentArray", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange CurrentRegion
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CurrentRegion", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Dependents 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Dependents", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange DirectDependents 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("DirectDependents", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange DirectPrecedents
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("DirectPrecedents", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange EntireColumn 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("EntireColumn", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange EntireRow 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("EntireRow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange End()
        {
            object returnValue = InstanceType.InvokeMember("End", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange End(XlDirection direction)
        {
            object[] paramArray = new object[1];
            paramArray[0] = direction;
            object returnValue = InstanceType.InvokeMember("End", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlCharacters Characters()
        {
            object returnValue = InstanceType.InvokeMember("Characters", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCharacters newClass = new XlCharacters(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlCharacters Characters(int start, int length)
        {
            object[] paramArray = new object[2];
            paramArray[0] = start;
            paramArray[1] = length;
            object returnValue = InstanceType.InvokeMember("Characters", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlCharacters newClass = new XlCharacters(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Item()
        {
            object returnValue = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Item(int rowIndex, int columnIndex)
        {
            object[] paramArray = new object[2];
            paramArray[0] = rowIndex;
            paramArray[1] = columnIndex;
            object returnValue = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Offset()
        {
            object returnValue = InstanceType.InvokeMember("Offset", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Offset(int rowOffset, int columnOffset)
        {
            object[] paramArray = new object[2];
            paramArray[0] = rowOffset;
            paramArray[1] = columnOffset;
            object returnValue = InstanceType.InvokeMember("Offset", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
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

        public XlRange MergeArea 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("MergeArea", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlSmartTags SmartTags 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("SmartTags", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlSmartTags newClass = new XlSmartTags(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWorksheet Worksheet 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Worksheet", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorksheet newClass = new XlWorksheet(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Range(XlRange cell1)
        {
            object[] paramArray = new object[1];
            paramArray[0] = cell1.COMReference;
            object returnValue = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Range(XlRange cell1, XlRange cell2)
        {
            object[] paramArray = new object[2];
            paramArray[0] = cell1.COMReference;
            paramArray[1] = cell2.COMReference;
            object returnValue = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Resize()
        {
            object returnValue = InstanceType.InvokeMember("Resize", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Resize(double rowSize, double columnSize)
        {
            object[] paramArray = new object[2];
            paramArray[0] = rowSize;
            paramArray[1] = columnSize;
            object returnValue = InstanceType.InvokeMember("Resize", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlPhonetic Phonetic 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Phonetic", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPhonetic newClass = new XlPhonetic(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlPhonetics Phonetics
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Phonetics", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPhonetics newClass = new XlPhonetics(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Precedents 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Precedents", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Previous 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Previous", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlQueryTable QueryTable 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Previous", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlQueryTable newClass = new XlQueryTable(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Next 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Next", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlComment Comment
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Comment", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlComment newClass = new XlComment(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlErrors Errors 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Errors", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlErrors newClass = new XlErrors(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlAreas Areas 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Areas", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlAreas newClass = new XlAreas(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlFormatConditions FormatConditions 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("FormatConditions", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlFormatConditions newClass = new XlFormatConditions(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        [System.Runtime.CompilerServices.IndexerNameAttribute("RangeItem")]
        public XlRange this[string adress]
        {
            get
            {
                
                object[] paramArray = new object[1];
                paramArray[0] = adress;
                object returnValue  = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        [System.Runtime.CompilerServices.IndexerNameAttribute("RangeItem")]
        public XlRange this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
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
            XlRange[] res_ranges = new XlRange[iCount];

            for (int i = 1; i <= iCount; i++)
                res_ranges[i - 1] = this[i];

            for (int i = 0; i < res_ranges.Length; i++)
            {
                yield return res_ranges[i];
            }

        }
        
        #endregion

        #region Methods

        public XlComment AddComment()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            object returnValue =InstanceType.InvokeMember("AddComment", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            XlComment newClass = new XlComment(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlComment AddComment(string text)
        {
            object[] paramArray = new object[1];
            paramArray[0] = text;
            object returnValue = InstanceType.InvokeMember("AddComment", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            XlComment newClass = new XlComment(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void BorderAround(XlLineStyle lineStyle, XlBorderWeight weight, double color)
        {
            object[] paramArray = new object[4];
            paramArray[0] = lineStyle;
            paramArray[1] = weight;
            paramArray[2] = Missing.Value;
            paramArray[3] = color;
            InstanceType.InvokeMember("BorderAround", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void BorderAround(XlLineStyle lineStyle, XlBorderWeight weight, XlColorIndex index)
        {
            object[] paramArray = new object[4];
            paramArray[0] = lineStyle;
            paramArray[1] = weight;
            paramArray[2] = index;
            paramArray[3] = Missing.Value;
            InstanceType.InvokeMember("BorderAround", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }


        public void BorderAround(XlLineStyle lineStyle, XlBorderWeight weight, XlColorIndex colorIndex, Color color)
        {
            object[] paramArray = new object[4];
            paramArray[0] = lineStyle;
            paramArray[1] = weight;
            paramArray[2] = colorIndex;
            paramArray[3] = XlConverter.ToDouble(color);
            InstanceType.InvokeMember("BorderAround", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void BorderAround(XlLineStyle lineStyle, XlBorderWeight weight)
        {
            object[] paramArray = new object[4];
            paramArray[0] = lineStyle;
            paramArray[1] = weight;
            paramArray[2] = Missing.Value;
            paramArray[3] = Missing.Value;
            InstanceType.InvokeMember("BorderAround", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void BorderAround(XlLineStyle lineStyle, XlBorderWeight weight, Color color)
        {
            object[] paramArray = new object[4];
            paramArray[0] = lineStyle;
            paramArray[1] = weight;
            paramArray[2] = Missing.Value;
            paramArray[3] = XlConverter.ToDouble(color);
            InstanceType.InvokeMember("BorderAround", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void AutoFit()
        {
            InstanceType.InvokeMember("AutoFit", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);           
        }

        public void Clear()
        {  
           InstanceType.InvokeMember("Clear", BindingFlags.InvokeMethod, null, ComReference, null,  XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PasteSpecial(XlPasteType pasteType)
        {
            object[] paramArray = new object[1];
            paramArray[0] = pasteType;
            InstanceType.InvokeMember("PasteSpecial", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Group()
        {
            object[] paramArray = new object[4];
            paramArray[0] = Missing.Value;
            paramArray[1] = Missing.Value;
            paramArray[2] = Missing.Value;
            paramArray[3] = Missing.Value;
            InstanceType.InvokeMember("Group", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Group(object start)
        {
            object[] paramArray = new object[4];
            paramArray[0] = start;
            paramArray[1] = Missing.Value;
            paramArray[2] = Missing.Value;
            paramArray[3] = Missing.Value;
            InstanceType.InvokeMember("Group", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Group(object start, object end)
        {
            object[] paramArray = new object[4];
            paramArray[0] = start;
            paramArray[1] = end;
            paramArray[2] = Missing.Value;
            paramArray[3] = Missing.Value;
            InstanceType.InvokeMember("Group", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Group(object start, object end, object by)
        {
            object[] paramArray = new object[4];
            paramArray[0] = start;
            paramArray[1] = end;
            paramArray[2] = by;
            paramArray[3] = Missing.Value;
            InstanceType.InvokeMember("Group", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Group(object start, object end, object by, object periods)
        {
            object[] paramArray = new object[4];
            paramArray[0] = start;
            paramArray[1] = end;
            paramArray[2] = by;
            paramArray[3] = periods;
            InstanceType.InvokeMember("Group", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Ungroup()
        {
            InstanceType.InvokeMember("Ungroup", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Insert()
        {
            object[] paramArray = new object[2];
            paramArray[0] = Missing.Value;
            paramArray[1] = Missing.Value;
            InstanceType.InvokeMember("Insert", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Insert(XlInsertShiftDirection shift)
        {
            object[] paramArray = new object[2];
            paramArray[0] = shift;
            paramArray[1] = Missing.Value;
            InstanceType.InvokeMember("Insert", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Insert(XlInsertShiftDirection shift, XlInsertFormatOrigin copyOrigin)
        {
            object[] paramArray = new object[2];
            paramArray[0] = shift;
            paramArray[1] = copyOrigin;
            InstanceType.InvokeMember("Insert", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Insert(XlInsertShiftDirection shift, XlRange copyOrigin)
        {
            object[] paramArray = new object[2];
            paramArray[0] = shift;
            paramArray[1] = copyOrigin.COMReference;
            InstanceType.InvokeMember("Insert", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void InsertIndent(int insertAmount)
        {
            object[] paramArray = new object[1];
            paramArray[0] = insertAmount;
            InstanceType.InvokeMember("InsertIndent", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Copy()
        {
            InstanceType.InvokeMember("Copy", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture); 
        }

        #endregion
    }
}
    