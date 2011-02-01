using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlCodeModule : XlNonCreatable
    {
        #region Construction

        internal XlCodeModule(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
		 
		}

        #endregion   
        
        #region Scalar Properties

        public int CountOfDeclarationLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CountOfDeclarationLines", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int CountOfLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CountOfLines", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int ProcBodyLine(string procName, vbext_ProcKind procKind)
        {
            object[] paramArray = new object[2];
            paramArray[0] = procName;
            paramArray[1] = procKind;
            object returnValue  = InstanceType.InvokeMember("ProcBodyLine", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (int)returnValue;
        }

        public int ProcCountLines(string procName, vbext_ProcKind procKind)
        {
            object[] paramArray = new object[2];
            paramArray[0] = procName;
            paramArray[1] = procKind;
            object returnValue  = InstanceType.InvokeMember("ProcCountLines", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (int)returnValue;
        }

        public int ProcOfLine(int line, vbext_ProcKind procKind)
        {
            object[] paramArray = new object[2];
            paramArray[0] = line;
            paramArray[1] = procKind;
            object returnValue  = InstanceType.InvokeMember("ProcOfLine", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (int)returnValue;
        }

        public int ProcStartLine(string procName, vbext_ProcKind procKind)
        {
            object[] paramArray = new object[2];
            paramArray[0] = procName;
            paramArray[1] = procKind;
            object returnValue  = InstanceType.InvokeMember("ProcStartLine", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (int)returnValue;
        }

        #endregion

        #region COMReference Properties

        public XlCodePane CodePane 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CodePane", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCodePane newClass = new XlCodePane(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBComponent Parent
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Parent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBComponent newClass = new XlVBComponent(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBE VBE
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBE", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBE newClass = new XlVBE(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Methods

        public void AddFromFile(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("AddFromFile", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);            
        }

        public void AddFromString(string String)
        {
            object[] paramArray = new object[1];
            paramArray[0] = String;
            InstanceType.InvokeMember("AddFromString", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public int CreateEventProc(string eventName, string objectName)
        {
            object[] paramArray = new object[2];
            paramArray[0] = eventName;
            paramArray[1] = objectName;
            object returnValue  = InstanceType.InvokeMember("CreateEventProc", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (int)returnValue;
        }

        public void DeleteLines(int startLine)
        {
            object[] paramArray = new object[1];
            paramArray[0] = startLine;
            InstanceType.InvokeMember("DeleteLines", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void DeleteLines(int startLine, int count)
        {
            object[] paramArray = new object[2];
            paramArray[0] = startLine;
            paramArray[1] = count;
            InstanceType.InvokeMember("DeleteLines", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public bool Find(string target, int startLine, int startColumn, int endLine, int endColumn)
        {
            object[] paramArray = new object[2];
            paramArray[0] = target;
            paramArray[1] = startLine;
            paramArray[2] = startColumn;
            paramArray[3] = endLine;
            paramArray[4] = endColumn;
            object returnValue  = InstanceType.InvokeMember("Find", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public void InsertLines(int line, string String)
        {
            object[] paramArray = new object[2];
            paramArray[0] = line;
            paramArray[1] = String;
            InstanceType.InvokeMember("InsertLines", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ReplaceLine(int Line, string String)
        {
            object[] paramArray = new object[2];
            paramArray[0] = Line;
            paramArray[1] = String;
            InstanceType.InvokeMember("ReplaceLine", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);

        }
        #endregion
    }
}
    