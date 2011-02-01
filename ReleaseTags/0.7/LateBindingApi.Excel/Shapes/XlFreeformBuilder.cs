using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Shapes
{
    public class XlFreeformBuilder : XlNonCreatable
    {
        #region Construction

        internal XlFreeformBuilder(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion 
        
        #region Methods

        public void AddNodes(MsoSegmentType segmentType, MsoEditingType editingType, Single x1, Single y1)
        {
            object[] paramArray = new object[4];
            paramArray[0] = segmentType;
            paramArray[1] = editingType;
            paramArray[2] = x1;
            paramArray[3] = y1;
            InstanceType.InvokeMember("AddNodes", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }
        
        public XlShape ConvertToShape()
        {
            object returnValue = InstanceType.InvokeMember("ConvertToShape", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlShape newClass = new XlShape(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
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

        #endregion

        #region Scalar Properties

        public XlCreator Creator
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        #endregion   
    }
}
