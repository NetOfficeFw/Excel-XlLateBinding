using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Stdole
{
    public class IPictureDisp : XlNonCreatable
    {
        #region Construction

        internal IPictureDisp(IXlObject parentReference, object comReference) : base(parentReference, comReference)
		{
        
		}

        #endregion

        #region Scalar Properties

        public int Handle 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Handle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Handle", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Height 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Height", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int hPal 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("hPal", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("hPal", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Type 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Type", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Type", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Width 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Width", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion

        #region Methods

        public void Render(int hdc,int x,int y,int cx,int cy ,int xSrc,int ySrc,int cxSrc ,int cySrc, object prcWBounds)
        {
            object[] paramArray = new object[10];
            paramArray[0] = hdc;
            paramArray[1] = x;
            paramArray[2] = y;
            paramArray[3] = cx;
            paramArray[4] = cy;
            paramArray[5] = xSrc;
            paramArray[6] = ySrc;
            paramArray[7] = cxSrc;
            paramArray[8] = cySrc;
            paramArray[9] = prcWBounds;
            InstanceType.InvokeMember("Render", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);           
        }

        #endregion
    }
}
