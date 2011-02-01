using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using XlLateBinding.Interfaces;

namespace XlLateBinding.Shapes
{
    public class XlHyperlink : XlNonCreatableBase
    {
        #region Construction

        internal XlHyperlink(IXlBase parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion

        public string Address
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("Address", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                _InstanceType.InvokeMember("Address", BindingFlags.SetProperty, null, _ComReference, paramArray, XlLateBindingSettings.XlThreadCulture);
            }
        }

        public void AddToFavorites()
        {
            _InstanceType.InvokeMember("AddToFavorites", BindingFlags.InvokeMethod, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
        }

        public void CreateNewDocument(string Filename, bool EditNow, bool Overwrite)
        {
            object[] paramArray = new object[3];
            paramArray[0] = Filename;
            paramArray[1] = EditNow;
            paramArray[2] = Overwrite;
            _InstanceType.InvokeMember("CreateNewDocument", BindingFlags.InvokeMethod, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
        }

        public void Delete()
        {
            _InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
        }

        public string EmailSubject
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("EmailSubject", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                _InstanceType.InvokeMember("EmailSubject", BindingFlags.SetProperty, null, _ComReference, paramArray, XlLateBindingSettings.XlThreadCulture);
            }
        }

        public string Name
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("Name", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                return (string)returnValue;
            }           
        }

        public XlRange Range
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                XlRange newClass = new XlRange(this, returnValue);
                _ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public string ScreenTip
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("ScreenTip", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                _InstanceType.InvokeMember("ScreenTip", BindingFlags.SetProperty, null, _ComReference, paramArray, XlLateBindingSettings.XlThreadCulture);
            }
        }


        public XlShape Shape 
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("Shape", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                XlShape newClass = new XlShape(this, returnValue);
                _ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public string SubAddress
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("SubAddress", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                _InstanceType.InvokeMember("SubAddress", BindingFlags.SetProperty, null, _ComReference, paramArray, XlLateBindingSettings.XlThreadCulture);
            }
        }


        public string TextToDisplay 
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("TextToDisplay", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                _InstanceType.InvokeMember("TextToDisplay", BindingFlags.SetProperty, null, _ComReference, paramArray, XlLateBindingSettings.XlThreadCulture);
            }
        }

        public int Type
        {
            get
            {
                object returnValue = _InstanceType.InvokeMember("Type", BindingFlags.GetProperty, null, _ComReference, null, XlLateBindingSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

    }
}
