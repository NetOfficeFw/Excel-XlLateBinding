using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Security;
using System.Security.Permissions;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Shapes;
using LateBindingApi.Excel.Office;
using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Pivot;
using LateBindingApi.Excel.Charts;

namespace LateBindingApi.Excel
{
    /// <summary>
    /// Represents a WorkSheet in Excel
    /// </summary>
    public class XlWorksheet : XlNonCreatable, IXlEventBinding
    {
        #region Fields

        XlWorksheetEvents _eventBridge;

        #endregion

        #region ComReference Properties

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

        public XlAutoFilter AutoFilter 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("AutoFilter", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlAutoFilter newClass = new XlAutoFilter(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
         
        public XlCustomProperties CustomProperties 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CustomProperties", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCustomProperties newClass = new XlCustomProperties(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlHPageBreaks HPageBreaks 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("HPageBreaks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlHPageBreaks newClass = new XlHPageBreaks(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVPageBreaks VPageBreaks
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("VPageBreaks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVPageBreaks newClass = new XlVPageBreaks(this, returnValue);
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

        public XlMsoEnvelope MailEnvelope 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("MailEnvelope", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlMsoEnvelope newClass = new XlMsoEnvelope(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlNames Names 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Names", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlNames newClass = new XlNames(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWorksheet Next 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Next", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorksheet newClass = new XlWorksheet(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWorksheet Previous 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Previous", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorksheet newClass = new XlWorksheet(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlProtection Protection 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Protection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlProtection newClass = new XlProtection(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlQueryTables QueryTables 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("QueryTables", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlQueryTables newClass = new XlQueryTables(this, returnValue);
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

        public XlSmartTags SmartTags 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("SmartTags", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlSmartTags newClass = new XlSmartTags(this, returnValue);
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

        public XlChartObjects ChartObjects
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ChartObjects", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlChartObjects newClass = new XlChartObjects(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlComments Comments
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Comments", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlComments newClass = new XlComments(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Cells(int rowIndex, int columnIndex)
        {
            object[] paramArray = new object[2];
            paramArray[0] = rowIndex;
            paramArray[1] = columnIndex;
            object returnValue  = InstanceType.InvokeMember("Cells", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange UsedRange
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UsedRange", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
 
        public XlRange Range(string adress)
        {
            object[] paramArray = new object[1];
            paramArray[0] = adress;
            object returnValue  = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange Range(string adress1, string adress2)
        {
            object[] paramArray = new object[2];
            paramArray[0] = adress1;
            paramArray[1] = adress2;
            object returnValue  = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            XlRange newClass = new XlRange(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public XlRange this[int rowIndex, int columnIndex]
        {
            get
            {
                object[] paramArray = new object[2];
                paramArray[0] = rowIndex;
                paramArray[1] = columnIndex;
                object returnValue  = InstanceType.InvokeMember("Cells", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;

            }
        }

        public XlRange this[string adress1]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = adress1;
                object returnValue = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;

            }
        }

        public XlRange this[string adress1, string adress]
        {
            get
            {
                object[] paramArray = new object[2];
                paramArray[0] = adress1;
                paramArray[1] = adress;
                object returnValue = InstanceType.InvokeMember("Range", BindingFlags.GetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;

            }
        }

        public XlRange Rows
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Rows", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange Columns
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Columns", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRange CircularReference 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CircularReference", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

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

        #endregion

        #region Construction

        internal XlWorksheet(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion

        #region Methods

        public void Delete()
        {
            InstanceType.InvokeMember("Delete", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }
 
        public void Select()
        {
            InstanceType.InvokeMember("Select", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PrintOut()
        {
            InstanceType.InvokeMember("PrintOut", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SaveAs(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("SaveAs", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Paste()
        {
            InstanceType.InvokeMember("Paste", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Paste(XlRange range)
        {
            object[] paramArray = new object[1];
            paramArray[0] = range.COMReference;
            InstanceType.InvokeMember("Paste", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Move(XlWorksheet before, XlWorksheet after)
        {
            object param1 = Missing.Value;
            object param2 = Missing.Value;

            if (null != before)
                param1 = before.COMReference;
            if (null != after)
                param2 = after.COMReference;

            object[] paramArray = new object[2];
            paramArray[0] = param1;
            paramArray[1] = param2;
            InstanceType.InvokeMember("Move", BindingFlags.InvokeMethod | BindingFlags.OptionalParamBinding, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion

        #region Events
        
        public event DocEvents_ActivateEventHandler Activate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseActivateEvent()
        {
            if (null == Activate) return;

            Activate();
        }

        public event DocEvents_BeforeDoubleClickEventHandler BeforeDoubleClick;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseBeforeDoubleClickEvent(object Target, ref bool Cancel)
        {
            if (null == BeforeDoubleClick)
            {
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlRange ra = new XlRange(this, Target);
            ListChildReferences.Add(ra);

            BeforeDoubleClick(ra, ref Cancel);
        }


        public event DocEvents_BeforeRightClickEventHandler BeforeRightClick;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseBeforeRightClickEvent(object Target, ref bool Cancel)
        {
            if (null == BeforeRightClick)
            {
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlRange ra = new XlRange(this, Target);
            ListChildReferences.Add(ra);

            BeforeRightClick(ra, ref Cancel);
        }

        public event DocEvents_CalculateEventHandler Calculate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseCalculateEvent()
        {
            if (null == Calculate) return;

            Calculate();
        }
        
        public event DocEvents_ChangeEventHandler Change;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseChangeEvent(object Target)
        {
            if (null == Change)
            {
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlRange ra = new XlRange(this, Target);
            ListChildReferences.Add(ra);

            Change(ra);
        }

        public event DocEvents_DeactivateEventHandler Deactivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseDeactivateEvent()
        {
            if (null == Deactivate) return;

            Deactivate();
        }

        public event DocEvents_FollowHyperlinkEventHandler FollowHyperlink;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseFollowHyperlinkEvent(object Target)
        {
            if (null == FollowHyperlink)
            {
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlHyperlink hl = new XlHyperlink(this, Target);
            ListChildReferences.Add(hl);

            FollowHyperlink(hl);
        }


        public event DocEvents_PivotTableUpdateEventHandler PivotTableUpdate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaisePivotTableUpdateEvent(object Target)
        {
            if (null == PivotTableUpdate)
            {
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlPivotTable pt = new XlPivotTable(this, Target);
            ListChildReferences.Add(pt);

            PivotTableUpdate(pt);
        }

        public event DocEvents_SelectionChangeEventHandler SelectionChange;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSelectionChangeEvent(object Target)
        {
            if (null == SelectionChange)
            {
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlRange ra = new XlRange(this, Target);
            ListChildReferences.Add(ra);

            SelectionChange(ra);
        }

        #endregion

        #region IXlEventBinding Members

        public void SetupEventBinding()
        {
            _eventBridge = new XlWorksheetEvents();
            _eventBridge.SetupEventBinding(this);
        }

        public void RemoveEventBinding()
        {
            _eventBridge.RemoveEventBinding();
        }
        
        #endregion
    }
}
