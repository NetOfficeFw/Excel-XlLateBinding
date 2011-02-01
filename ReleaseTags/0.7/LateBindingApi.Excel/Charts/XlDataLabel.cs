using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Office;
using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Pivot;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Charts
{
    public class XlDataLabel : XlNonCreatable
    { 
        #region Construction

        internal XlDataLabel(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		}

        #endregion

        #region Methods

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);

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

        public XlChartGroup Bar3DGroup 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Bar3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartArea ChartArea 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ChartArea", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartArea newClass = new XlChartArea(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartTitle ChartTitle
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ChartTitle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartTitle newClass = new XlChartTitle(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartGroup Column3DGroup
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Column3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCorners Corners
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Corners", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCorners newClass = new XlCorners(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDataTable DataTable 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("DataTable", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDataTable newClass = new XlDataTable(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlFloor Floor 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Floor", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlFloor newClass = new XlFloor(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlHyperlinks Hyperlinks 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Hyperlinks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlHyperlinks newClass = new XlHyperlinks(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlLegend Legend 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Legend", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlLegend newClass = new XlLegend(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartGroup Line3DGroup 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Line3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlPageSetup PageSetup
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("PageSetup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPageSetup newClass = new XlPageSetup(this, returnValue);
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

        public XlChartGroup Pie3DGroup 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Pie3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlPivotLayout PivotLayout 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("PivotLayout", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPivotLayout newClass = new XlPivotLayout(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlPlotArea PlotArea 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("PlotArea", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPlotArea newClass = new XlPlotArea(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlNonCreatable Previous 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Previous", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlNonCreatable newClass = XlDynamicType.CreateDynamicType(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlScripts Scripts
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Scripts", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlScripts newClass = new XlScripts(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlSeriesCollection SeriesCollection()
        {
            object returnValue = InstanceType.InvokeMember("SeriesCollection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlSeriesCollection newClass = new XlSeriesCollection(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlSeries SeriesCollection(int index)
        {
            object[] parameter = new object[1];
            parameter[0] = index;
            object returnValue = InstanceType.InvokeMember("SeriesCollection", BindingFlags.GetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlSeries newClass = new XlSeries(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlShapes Shapes 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Shapes", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShapes newClass = new XlShapes(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartGroup SurfaceGroup 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("SurfaceGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlTab Tab
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Tab", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlTab newClass = new XlTab(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWalls Walls
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Walls", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWalls newClass = new XlWalls(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartGroups  XYGroups()
        {
            object returnValue = InstanceType.InvokeMember("XYGroups", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroup XYGroups(int index)
        {
            object[] parameter = new object[1];
            parameter[0] = index;
            object returnValue = InstanceType.InvokeMember("XYGroups", BindingFlags.GetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;

        }

        public XlChartGroups PieGroups()
        {
            object returnValue = InstanceType.InvokeMember("PieGroups", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroup PieGroups(int index)
        {
            object[] parameter = new object[1];
            parameter[0] = index;
            object returnValue = InstanceType.InvokeMember("PieGroups", BindingFlags.GetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

 

        #endregion

        #region Scalar Properties

        public bool AutoScaling 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AutoScaling", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AutoScaling", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlBarShape BarShape 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("BarShape", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlBarShape)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("BarShape", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlChartType ChartType 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ChartType", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlChartType)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ChartType", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string CodeName
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CodeName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public XlCreator Creator
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public int DepthPercent 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("DepthPercent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public XlDisplayBlanksAs DisplayBlanksAs 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("DisplayBlanksAs", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlDisplayBlanksAs)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayBlanksAs", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Elevation 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Elevation", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Elevation", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int GapDepth 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("GapDepth", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("GapDepth", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasAxis
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HasAxis", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool HasDataTable
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HasDataTable", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool HasLegend 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HasLegend", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool HasPivotFields 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HasPivotFields", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool HasTitle 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HasTitle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public int HeightPercent 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HeightPercent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("HeightPercent", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Index  
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Index", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public XlMsoEnvelope MailEnvelope 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("MailEnvelope", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlMsoEnvelope)returnValue;
            }
        }

        public string Name
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Name", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Name", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Perspective 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Perspective", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public XlRowCol PlotBy 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("PlotBy", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlRowCol)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("PlotBy", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool PlotVisibleOnly 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("PlotVisibleOnly", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("PlotVisibleOnly", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectContents 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ProtectContents", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ProtectData
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ProtectData", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ProtectData", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectDrawingObjects 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ProtectDrawingObjects", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ProtectFormatting 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ProtectFormatting", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ProtectFormatting", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectGoalSeek 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ProtectGoalSeek", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ProtectGoalSeek", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectionMode 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ProtectionMode ", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ProtectSelection 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ProtectSelection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ProtectSelection", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowWindow 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("ShowWindow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ShowWindow", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool SizeWithWindow 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("SizeWithWindow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("SizeWithWindow", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlSheetVisibility Visible 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlSheetVisibility)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool WallsAndGridlines2D 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("WallsAndGridlines2D", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("WallsAndGridlines2D", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
    }
}
