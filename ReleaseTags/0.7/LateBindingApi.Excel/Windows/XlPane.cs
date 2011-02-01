using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Windows
{
    public class XlPane: XlNonCreatable 
    {
        #region Construction

        internal XlPane(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{		 

		}

        #endregion       

        #region COMReference Properties

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

        public XlRange VisibleRange
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VisibleRange", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        
        #endregion

        #region Scalar Properties
 
        public XlCreator Creator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public int Index 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Index", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int ScrollColumn 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ScrollColumn", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int ScrollRow 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ScrollRow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion

        #region Methods

        public bool Activate()
        {
            object returnValue  = InstanceType.InvokeMember("Activate", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public bool LargeScroll()
        {
            object returnValue  = InstanceType.InvokeMember("LargeScroll", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public bool SmallScroll()
        {
            object returnValue  = InstanceType.InvokeMember("SmallScroll", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public bool ScrollIntoView(int left, int top, int width, int height)
        {
            object[] paramArray = new object[4];
            paramArray[0] = left;
            paramArray[1] = top;
            paramArray[2] = width;
            paramArray[3] = height;
            object returnValue  = InstanceType.InvokeMember("ScrollIntoView", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        #endregion
    }
}
