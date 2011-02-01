using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Charts
{
    public class XlHiLoLines : XlNonCreatable 
    {   
        #region Construction

        internal XlHiLoLines(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{

		}

        #endregion

        #region COMReference Properties

        public XlDownBars DownBars
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DownBars", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;          
                XlDownBars newClass = new XlDownBars(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDropLines DropLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DropLines", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;  
                XlDropLines newClass = new XlDropLines(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlHiLoLines HiLoLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HiLoLines", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;  
                XlHiLoLines newClass = new XlHiLoLines(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlTickLabels RadarAxisLabels
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RadarAxisLabels", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;  
                XlTickLabels newClass = new XlTickLabels(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlSeriesLines SeriesLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SeriesLines", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;  
                XlSeriesLines newClass = new XlSeriesLines(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlUpBars UpBars
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UpBars", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;  
                XlUpBars newClass = new XlUpBars(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public XlAxisGroup AxisGroup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AxisGroup", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlAxisGroup)returnValue;
            }
        }

        public int BubbleScale
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BubbleScale", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int DoughnutHoleSize
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DoughnutHoleSize", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("DoughnutHoleSize", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int FirstSliceAngle
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FirstSliceAngle", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("FirstSliceAngle", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int GapWidth
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GapWidth", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("GapWidth", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Has3DShading
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Has3DShading", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Has3DShading", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasDropLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasDropLines", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasDropLines", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasHiLoLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasHiLoLines", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasHiLoLines", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasRadarAxisLabels
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasRadarAxisLabels", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasRadarAxisLabels", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasSeriesLines
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasSeriesLines", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasSeriesLines", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasUpDownBars
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasUpDownBars", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasUpDownBars", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
        public int Index
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Index", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int Overlap
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Overlap", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Overlap", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int SecondPlotSize
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SecondPlotSize", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("SecondPlotSize", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
        public bool ShowNegativeBubbles
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowNegativeBubbles", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ShowNegativeBubbles", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlSizeRepresents SizeRepresents
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SizeRepresents", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlSizeRepresents)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("SizeRepresents", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlChartSplitType SplitType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SplitType", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlChartSplitType)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("SplitType", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool VaryByCategories
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VaryByCategories", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("VaryByCategories", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }
}
