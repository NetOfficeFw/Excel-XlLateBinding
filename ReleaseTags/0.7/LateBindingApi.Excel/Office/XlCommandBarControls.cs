using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Office
{
    public class XlCommandBarControls : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlCommandBarControls(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion  

        #region Methods

        public XlCommandBarControl Add(MsoControlType type)
        {
            object[] paramArray = new object[5];
            paramArray[0] = type;
            paramArray[1] = Missing.Value;
            paramArray[2] = Missing.Value;
            paramArray[3] = Missing.Value;
            paramArray[4] = Missing.Value;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            return ReturnHelper(type, returnValue);
        }

        public XlCommandBarControl Add(MsoControlType type, object id, object parameter, object before, bool temporary)
        {
            object[] paramArray = new object[5];
            paramArray[0] = type;
            paramArray[1] = id;
            paramArray[2] = parameter;
            paramArray[3] = before;
            paramArray[4] = temporary;
            object returnValue  = InstanceType.InvokeMember("Add", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            return ReturnHelper(type, returnValue);
        }

        private XlCommandBarControl ReturnHelper(MsoControlType type, object comRef)
        {
            switch (type)
            {
                case MsoControlType.msoControlButton:

                    XlCommandBarButton newButton = new XlCommandBarButton(this, comRef);
                    ListChildReferences.Add(newButton);
                    return newButton;

                case MsoControlType.msoControlPopup:

                    XlCommandBarPopup newPopup = new XlCommandBarPopup(this, comRef);
                    ListChildReferences.Add(newPopup);
                    return newPopup;

                case MsoControlType.msoControlComboBox:

                    XlCommandBarComboBox newBox = new XlCommandBarComboBox(this, comRef);
                    ListChildReferences.Add(newBox);
                    return newBox;

                default:

                    XlCommandBarControl newClass = new XlCommandBarControl(this, comRef);
                    ListChildReferences.Add(newClass);
                    return newClass;
            }

        }

        #endregion

        #region COMReference Properties

        /// <summary>
        /// returns an XlCommandBarControl by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlCommandBarControl this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object comRef  = InstanceType.InvokeMember("Item", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                XlCommandBarControl newClass = new XlCommandBarControl(this, comRef);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

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

        public int Creator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
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
            XlCommandBarControl[] res_addins = new XlCommandBarControl[iCount];

            for (int i = 1; i <= iCount; i++)
                res_addins[i - 1] = this[i];

            for (int i = 0; i < res_addins.Length; i++)
            {
                yield return res_addins[i];
            }

        }

        #endregion
    }
}
