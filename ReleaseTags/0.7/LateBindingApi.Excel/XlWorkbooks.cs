using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel
{
    public class XlWorkbooks : XlNonCreatable, System.Collections.IEnumerable
    {  
        #region Construction

        internal XlWorkbooks(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
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
            XlWorkbook[] res_books = new XlWorkbook[iCount];

            for (int i = 1; i <= iCount; i++)
                res_books[i - 1] = this[i];

            for (int i = 0; i < res_books.Length; i++)
            {
                yield return res_books[i];
            }

        }

        #endregion
  
        #region Scalar Properties

        /// <summary>
        /// Count of Workbooks
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
        /// returns an Workbook by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlWorkbook this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorkbook newClass = new XlWorkbook(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        
        #endregion

        #region Methods

        public bool CanCheckOut(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            object returnValue = InstanceType.InvokeMember("CanCheckOut", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue; 
        }

        public void CheckOut(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("CheckOut", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlWorkbook OpenXML(String filename)
        {
            object[] paramArray = new object[1];
            paramArray[0] = filename;

            object returnValue = InstanceType.InvokeMember("OpenXML", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;

            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlWorkbook OpenXML(String filename, object styleSheets)
        {
            object[] paramArray = new object[2];
            paramArray[0] = filename;
            paramArray[1] = styleSheets;

            object returnValue = InstanceType.InvokeMember("OpenXML", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;

            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlWorkbook OpenDatabase(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            object returnValue = InstanceType.InvokeMember("OpenDatabase", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlWorkbook OpenDatabase(string fileName, object commandText, object commandType, object backgroundQuery, object importDataAs)
        {
            object[] paramArray = new object[5];
            paramArray[0] = fileName;
            paramArray[1] = commandText;
            paramArray[2] = commandType;
            paramArray[3] = backgroundQuery;
            paramArray[4] = importDataAs;
            object returnValue = InstanceType.InvokeMember("OpenDatabase", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlWorkbook OpenText(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            object returnValue = InstanceType.InvokeMember("OpenText", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void OpenText(String Filename, object Origin, object StartRow, object DataType, XlTextQualifier TextQualifier, object ConsecutiveDelimiter, object Tab, object Semicolon, object Comma, object Space, object Other, object OtherChar, object FieldInfo, object TextVisualLayout, object DecimalSeparator, object ThousandsSeparator, object TrailingMinusNumbers, object Local)
        {
            object[] paramArray = new object[18];
            paramArray[0] = Filename;
            paramArray[1] = Origin;
            paramArray[2] = StartRow;
            paramArray[3] = DataType;
            paramArray[4] = TextQualifier;
            paramArray[5] = ConsecutiveDelimiter;
            paramArray[6] = Tab;
            paramArray[7] = Semicolon;
            paramArray[8] = Comma;
            paramArray[9] = Space;
            paramArray[10] = Other;
            paramArray[11] = OtherChar;
            paramArray[12] = FieldInfo;
            paramArray[13] = TextVisualLayout;
            paramArray[14] = DecimalSeparator;
            paramArray[15] = ThousandsSeparator;
            paramArray[16] = TrailingMinusNumbers;
            paramArray[17] = Local;

            InstanceType.InvokeMember("OpenText", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }


        /// <summary>
        /// Loads an Excelfile on local Filesystem
        /// </summary>
        /// <param name="Filename"></param>
        /// <returns></returns>
        public XlWorkbook Open(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            object returnValue = InstanceType.InvokeMember("Open", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlWorkbook Open(string fileName,
            object updateLinks, bool readOnly, object format, object passWord, object writeResPassword,
            object ignoreReadOnlyRecommended, object origin, object delimiter, object editable, object notify, object converter, object addToMru, object local, object corruptLoad)
        {
            object[] paramArray = new object[15];
            paramArray[0] = fileName;

            paramArray[1] = updateLinks;
            paramArray[2] = readOnly;
            paramArray[3] = format;
            paramArray[4] = passWord;
            paramArray[5] = writeResPassword;
            paramArray[6] = ignoreReadOnlyRecommended;
            paramArray[7] = origin;
            paramArray[8] = delimiter;
            paramArray[9] = editable;
            paramArray[10] = notify;
            paramArray[11] = converter;
            paramArray[12] = addToMru;
            paramArray[13] = local;
            paramArray[14] = corruptLoad;

            object returnValue = InstanceType.InvokeMember("Open", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        /// <summary>
        /// Append a new Workbook
        /// </summary>
        /// <returns></returns>
        public XlWorkbook Add()
        {
            object returnValue = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlWorkbook newClass = new XlWorkbook(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        #endregion
    }
}
