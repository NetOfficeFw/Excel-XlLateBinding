using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Office;
using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Pivot;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.Charts
{   
    /// <summary>
    /// Represents a XlChart in Excel
    /// </summary>
    public class XlChart : XlNonCreatable , IXlEventBinding
    {
        #region Fields
        
        XlChartEvents _eventBridge;

        #endregion

        #region Construction

        internal XlChart(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{           

		}

        #endregion

        #region Methods

        public void Activate()
        {
            InstanceType.InvokeMember("Activate", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ApplyCustomType(XlChartType chartType)
        {
            object[] paramArray = new object[1];
            paramArray[0] = chartType;
            InstanceType.InvokeMember("ApplyCustomType", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlChartGroups AreaGroups()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("AreaGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroup AreaGroups(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue = InstanceType.InvokeMember("AreaGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlAxes Axes()
        {
            object[] paramArray = new object[2];
            paramArray[0] = Missing.Value;
            paramArray[1] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("AreaGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlAxes newClass = new XlAxes(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlAxis Axes(object type, XlAxisGroup axisGroup)
        {
            object[] paramArray = new object[2];
            paramArray[0] = type;
            paramArray[1] = axisGroup;
            object returnValue = InstanceType.InvokeMember("AreaGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlAxis newClass = new XlAxis(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroups BarGroups()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("BarGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroup BarGroups(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue = InstanceType.InvokeMember("BarGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroups ChartGroups()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("ChartGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroup ChartGroups(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue = InstanceType.InvokeMember("ChartGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartObjects ChartObjects()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("ChartObjects", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartObjects newClass = new XlChartObjects(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartObject ChartObjects(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue = InstanceType.InvokeMember("ChartObjects", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartObject newClass = new XlChartObject(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroup ColumnGroups()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("ChartObjects", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroups ColumnGroups(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue = InstanceType.InvokeMember("ChartObjects", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Copy()
        {
            object[] paramArray = new object[2];
            paramArray[0] = Missing.Value;
            paramArray[1] = Missing.Value;
            InstanceType.InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Copy(XlNonCreatable before, XlNonCreatable after)
        {
            object[] paramArray = new object[2];
            paramArray[0] = before.ComReference;
            paramArray[1] = after.ComReference;
            InstanceType.InvokeMember("Copy", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void CopyPicture()
        {
            object[] paramArray = new object[3];
            paramArray[0] = XlPictureAppearance.xlScreen;
            paramArray[1] = XlCopyPictureFormat.xlPicture;
            paramArray[2] = XlPictureAppearance.xlPrinter;
            InstanceType.InvokeMember("CopyPicture", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void CopyPicture(XlPictureAppearance appearance, XlCopyPictureFormat format, XlPictureAppearance size)
        {
            object[] paramArray = new object[3];
            paramArray[0] = appearance;
            paramArray[1] = format;
            paramArray[2] = size;
            InstanceType.InvokeMember("CopyPicture", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Delete()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            paramArray[1] = Missing.Value;
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Deselect()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            paramArray[1] = Missing.Value;
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlChartGroup DoughnutGroups()
        {
            object[] paramArray = new object[1];
            paramArray[0] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("DoughnutGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroups DoughnutGroups(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            object returnValue = InstanceType.InvokeMember("DoughnutGroups", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Evaluate(string name)
        {
            object[] paramArray = new object[1];
            paramArray[0] = name;
            InstanceType.InvokeMember("Evaluate", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Export(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("Export", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void GetChartElement(int x, int y, int elemendId, int arg1, int arg2)
        {
            object[] paramArray = new object[5];
            paramArray[0] = x;
            paramArray[1] = y;
            paramArray[2] = elemendId;
            paramArray[3] = arg1;
            paramArray[4] = arg2;
            InstanceType.InvokeMember("GetChartElement", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlChart Location(XlChartLocation where)
        {
            object[] paramArray = new object[2];
            paramArray[0] = where;
            paramArray[1] = Missing.Value;
            object returnValue = InstanceType.InvokeMember("Location", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChart newClass = new XlChart(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }


        public void Move(XlNonCreatable before, XlNonCreatable after)
        {
            object[] paramArray = new object[2];
            paramArray[0] = before.ComReference;
            paramArray[1] = after.ComReference;
            InstanceType.InvokeMember("Move", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlOLEObjects OLEObjects()
        {
            object returnValue = InstanceType.InvokeMember("OLEObjects", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlOLEObjects newClass = new XlOLEObjects(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlOLEObject OLEObjects(int index)
        {
            object[] parameter = new object[1];
            parameter[0] = index;
            object returnValue = InstanceType.InvokeMember("OLEObjects", BindingFlags.GetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlOLEObject newClass = new XlOLEObject(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void Paste()
        {
            InstanceType.InvokeMember("Paste", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PrintOut()
        {
            InstanceType.InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PrintPreview(bool enableChanges)
        {
            object[] parameter = new object[1];
            parameter[0] = enableChanges;
            InstanceType.InvokeMember("PrintPreview", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Protect(string password)
        {
            object[] parameter = new object[1];
            parameter[0] = password;
            InstanceType.InvokeMember("Protect", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Refresh()
        {
            InstanceType.InvokeMember("Refresh", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public XlChartGroups RadarGroups()
        {
            object returnValue = InstanceType.InvokeMember("RadarGroups", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroups newClass = new XlChartGroups(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlChartGroup RadarGroups(int index)
        {
            object[] parameter = new object[1];
            parameter[0] = index;
            object returnValue = InstanceType.InvokeMember("RadarGroups", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlChartGroup newClass = new XlChartGroup(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void SaveAs(string fileName)
        {
            object[] parameter = new object[1];
            parameter[0] = fileName;
            InstanceType.InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetBackgroundPicture(string fileName)
        {
            object[] parameter = new object[1];
            parameter[0] = fileName;
            InstanceType.InvokeMember("SetBackgroundPicture", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetSourceData(XlRange source)
        {
            object[] parameter = new object[1];
            parameter[0] = source.COMReference;
            InstanceType.InvokeMember("SetSourceData", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Unprotect(string password)
        {
            object[] parameter = new object[1];
            parameter[0] = password;
            InstanceType.InvokeMember("Unprotect", BindingFlags.InvokeMethod, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region COMReference Properties

        public XlChartGroup Area3DGroup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Area3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartGroup Bar3DGroup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Bar3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartTitle ChartTitle
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ChartTitle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartTitle newClass = new XlChartTitle(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlChartArea ChartArea
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ChartArea", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartArea newClass = new XlChartArea(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlChartGroup Column3DGroup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Column3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDataTable DataTable
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DataTable", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("Floor", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("Hyperlinks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("Legend", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("Line3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartGroup newClass = new XlChartGroup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlMsoEnvelope MailEnvelope
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MailEnvelope", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlMsoEnvelope newClass = new XlMsoEnvelope(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlPageSetup PageSetup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PageSetup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPageSetup newClass = new XlPageSetup(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlChartGroup Pie3DGroup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Pie3DGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("PivotLayout", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("PlotArea", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPlotArea newClass = new XlPlotArea(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWalls Walls
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Walls", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWalls newClass = new XlWalls(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlScripts Scripts
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Scripts", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlScripts newClass = new XlScripts(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlShapes Shapes
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Shapes", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("SurfaceGroup", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
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
                object returnValue  = InstanceType.InvokeMember("Tab", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlTab newClass = new XlTab(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public bool AutoScaling
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoScaling", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoScaling", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlBarShape BarShape
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("BarShape", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlBarShape)returnValue;
            }
        }

        public XlChartType ChartType
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ChartType", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlChartType)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ChartType", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string CodeName
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CodeName", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public int DepthPercent
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DepthPercent", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("DepthPercent", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlDisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayBlanksAs", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlDisplayBlanksAs)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("DisplayBlanksAs", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int GapDepth
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GapDepth", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("GapDepth", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasDataTable
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasDataTable", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasDataTable", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasLegend
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasLegend", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasLegend", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasPivotFields
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasPivotFields", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasPivotFields", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasTitle
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasTitle", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasTitle", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int HeightPercent
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HeightPercent", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HeightPercent", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Name", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
 
        public int Perspective
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Perspective", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Perspective", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
 
        public XlRowCol PlotBy
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PlotBy", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlRowCol)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("PlotBy", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool PlotVisibleOnly
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PlotVisibleOnly", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("PlotVisibleOnly", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectContents
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectContents", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ProtectData
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectData", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ProtectData", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectDrawingObjects
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectDrawingObjects", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ProtectFormatting
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectFormatting", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ProtectFormatting", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectGoalSeek
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectGoalSeek", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ProtectGoalSeek", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ProtectionMode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectionMode", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ProtectSelection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectSelection", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ProtectSelection", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowWindow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowWindow", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ShowWindow", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool SizeWithWindow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SizeWithWindow", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("SizeWithWindow", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlSheetVisibility Visible
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlSheetVisibility)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool WallsAndGridlines2D
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WallsAndGridlines2D", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("WallsAndGridlines2D", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
        
        #region Events
        
        /// <summary>
        /// Event Activate mapped to ActivateEvent. Reason: Chart has an Activate Method
        /// </summary>
        public event ChartEvents_ActivateEventHandler ActivateEvent;
        internal void RaiseActivateEvent()
        {
            ActivateEvent();
        }

        public event ChartEvents_BeforeDoubleClickEventHandler BeforeDoubleClick;
        internal void RaiseBeforeDoubleClickEvent(int ElementID, int Arg1, int Arg2, ref bool Cancel)
        {
            BeforeDoubleClick(ElementID, Arg1, Arg2, ref Cancel);
        }

        public event ChartEvents_BeforeRightClickEventHandler BeforeRightClick;
        internal void RaiseBeforeRightClickEvent(ref bool Cancel)
        {
            BeforeRightClick(ref Cancel);
        }

        public event ChartEvents_CalculateEventHandler Calculate;
        internal void RaiseCalculateEvent()
        {
            Calculate();
        }

        public event ChartEvents_DeactivateEventHandler Deactivate;
        internal void RaiseDeactivateEvent()
        {
            Deactivate();
        }

        public event ChartEvents_DragOverEventHandler DragOver;
        internal void RaiseDragOverEvent()
        {
            DragOver();
        }

        public event ChartEvents_DragPlotEventHandler DragPlot;
        internal void RaiseDragPlotEvent()
        {
            DragPlot();
        }

        public event ChartEvents_MouseDownEventHandler MouseDown;
        internal void RaiseMouseDownEvent(int Button, int Shift, int x, int y)
        {
            MouseDown(Button, Shift,x,y);
        }

        public event ChartEvents_MouseMoveEventHandler MouseMove;
        internal void RaiseMouseMoveEvent(int Button, int Shift, int x, int y)
        {
            MouseMove(Button, Shift, x, y);
        }

        public event ChartEvents_MouseUpEventHandler MouseUp;
        internal void RaiseMouseUpEvent(int Button, int Shift, int x, int y)
        {
            MouseUp(Button, Shift, x, y);
        }

        public event ChartEvents_ResizeEventHandler Resize;
        internal void RaiseResizeEvent()
        {
            Resize();
        }

        public event ChartEvents_SelectEventHandler SelectEvent;
        internal void RaiseSelectEvent(int ElementID, int Arg1, int Arg2)
        {
            SelectEvent(ElementID, Arg1, Arg2);
        }
        
        public event ChartEvents_SeriesChangeEventHandler SeriesChange;
        internal void RaiseSeriesChangeEvent(int SeriesIndex, int PointIndex)
        {
            SeriesChange(SeriesIndex, PointIndex);
        }

        #endregion

        #region IXlEventBinding Members

        public void SetupEventBinding()
        {
            _eventBridge = new XlChartEvents();
            _eventBridge.SetupEventBinding(this);
        }

        public void RemoveEventBinding()
        {
            _eventBridge.RemoveEventBinding();
        }

        #endregion
    }
}
