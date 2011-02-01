using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using LateBindingApi.Excel.Windows;
using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Pivot;

namespace LateBindingApi.Excel
{
    #region Event Delegates

    public delegate void AppEvents_NewWorkbookEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb);
    public delegate void AppEvents_SheetActivateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] XlWorksheet Sh);
    public delegate void AppEvents_SheetBeforeDoubleClickEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] XlWorksheet Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target, [In] ref bool Cancel);
    public delegate void AppEvents_SheetBeforeRightClickEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] XlWorksheet Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target, [In] ref bool Cancel);
    public delegate void AppEvents_SheetCalculateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] XlWorksheet Sh);
    public delegate void AppEvents_SheetChangeEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] XlWorksheet Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target);
    public delegate void AppEvents_SheetDeactivateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] XlWorksheet Sh);
    public delegate void AppEvents_SheetFollowHyperlinkEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlHyperlink Target);
    public delegate void AppEvents_SheetPivotTableUpdateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlPivotTable Target);
    public delegate void AppEvents_SheetSelectionChangeEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target);
    public delegate void AppEvents_WindowActivateEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In, MarshalAs(UnmanagedType.Interface)] XlWindow Wn);   
    public delegate void AppEvents_WindowDeactivateEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In, MarshalAs(UnmanagedType.Interface)] XlWindow Wn);
    public delegate void AppEvents_WindowResizeEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In, MarshalAs(UnmanagedType.Interface)] XlWindow Wn);
    public delegate void AppEvents_WorkbookActivateEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb);
    public delegate void AppEvents_WorkbookAddinInstallEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb);
    public delegate void AppEvents_WorkbookAddinUninstallEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb);
    public delegate void AppEvents_WorkbookBeforeCloseEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In] ref bool Cancel);
    public delegate void AppEvents_WorkbookBeforePrintEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In] ref bool Cancel);
    public delegate void AppEvents_WorkbookBeforeSaveEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In] bool SaveAsUI, [In] ref bool Cancel);
    public delegate void AppEvents_WorkbookDeactivateEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb);
    public delegate void AppEvents_WorkbookNewSheetEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
    public delegate void AppEvents_WorkbookOpenEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb);
    public delegate void AppEvents_WorkbookPivotTableCloseConnectionEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In, MarshalAs(UnmanagedType.Interface)] XlPivotTable Target);
    public delegate void AppEvents_WorkbookPivotTableOpenConnectionEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWorkbook Wb, [In, MarshalAs(UnmanagedType.Interface)] XlPivotTable Target);

    #endregion
 
    [ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface IAppEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61d)]
        void NewWorkbook([In, MarshalAs(UnmanagedType.Interface)] object wb);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x616)]
        void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x617)]
        void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x618)]
        void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x619)]
        void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61a)]
        void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61b)]
        void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61c)]
        void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61f)]
        void WorkbookOpen([In, MarshalAs(UnmanagedType.Interface)] object wb);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x620)]
        void WorkbookActivate([In, MarshalAs(UnmanagedType.Interface)] object wb);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x621)]
        void WorkbookDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wb);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x622)]
        void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x623)]
        void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] bool SaveAsUI, [In] ref bool cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x624)]
        void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x625)]
        void WorkbookNewSheet([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x626)]
        void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.Interface)] object wb);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x627)]
        void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.Interface)] object wb);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x612)]
        void WindowResize([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x614)]
        void WindowActivate([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x615)]
        void WindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x73e)]
        void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x86d)]
        void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x870)]
        void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x871)]
        void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object target);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class XlApplicationEvents : IAppEvents
    {
        #region Fields

        private XlApplication _application;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion
        
        #region Construction

        public XlApplicationEvents()
        {
        }
        
        #endregion

        #region Events
  
        public void NewWorkbook(object __p1)
        {
            _application.RaiseNewWorkbookEvent(__p1); 
        }

        public void SheetActivate(object __p1)
        {
            _application.RaiseSheetActivateEvent(__p1);
        }

        public void SheetBeforeDoubleClick(object __p1, object __p2, ref bool __p3) 
        {
            _application.RaiseSheetBeforeDoubleClickEvent(__p1, __p2, ref __p3);
        }

        public void SheetBeforeRightClick(object __p1, object __p2, ref bool __p3) 
        {
            _application.RaiseSheetBeforeRightClickEvent(__p1, __p2, ref __p3);
        }

        public void SheetCalculate(object __p1) 
        {
            _application.RaiseSheetCalculateEvent(__p1);
        }

        public void SheetChange(object __p1, object __p2)
        {
            _application.RaiseSheetChangeEvent(__p1, __p2);
        }

        public void SheetDeactivate(object __p1) 
        {
            _application.RaiseSheetDeactivateEvent(__p1);
        }

        public void SheetFollowHyperlink(object __p1, object __p2) 
        {
            _application.RaiseSheetFollowHyperlinkEvent(__p1, __p2);
        }

        public void SheetPivotTableUpdate(object __p1, object __p2) 
        {
            _application.RaiseSheetPivotTableUpdateEvent(__p1, __p2);
        }

        public void SheetSelectionChange(object __p1, object __p2) 
        {
            _application.RaiseSheetSelectionChangeEvent(__p1, __p2);
        }

        public void WindowActivate(object __p1, object __p2)
        {
            _application.RaiseWindowActivateEvent(__p1, __p2);
        }

        public void WindowDeactivate(object __p1, object __p2) 
        {
            _application.RaiseWindowDeactivateEvent(__p1, __p2);
        }

        public void WindowResize(object __p1, object __p2) 
        {
            _application.RaiseWindowResizeEvent(__p1, __p2);
        }

        public void WorkbookActivate(object __p1) 
        {
            _application.RaiseWorkbookActivateEvent(__p1);
        }

        public void WorkbookAddinInstall(object __p1) 
        {
            _application.RaiseWorkbookAddinInstallEvent(__p1);
        }

        public void WorkbookAddinUninstall(object __p1) 
        {
            _application.RaiseWorkbookAddinUninstallEvent(__p1);
        }

        public void WorkbookBeforeClose(object __p1, ref bool __p2)
        {
            _application.RaiseWorkbookBeforeCloseEvent(__p1, ref __p2);
        }

        public void WorkbookBeforePrint(object __p1, ref bool __p2) 
        {
            _application.RaiseWorkbookBeforePrintEvent(__p1, ref __p2);
        }

        public void WorkbookBeforeSave(object __p1, bool __p2, ref bool __p3) 
        {
            _application.RaiseWorkbookBeforeSaveEvent(__p1, __p2, ref __p3);
        }

        public void WorkbookDeactivate(object __p1) 
        {
            _application.RaiseWorkbookDeactivateEvent(__p1);
        }

        public void WorkbookNewSheet(object __p1, object __p2) 
        {
            _application.RaiseWorkbookNewSheetEvent(__p1, __p2);
        }

        public void WorkbookOpen(object __p1)
        {
            _application.RaiseWorkbookOpenEvent(__p1);
        }

        public void WorkbookPivotTableCloseConnection(object __p1, object __p2) 
        {
            _application.RaiseWorkbookPivotTableCloseConnectionEvent(__p1, __p2);
        }

        public void WorkbookPivotTableOpenConnection(object __p1, object __p2) 
        {
            _application.RaiseWorkbookPivotTableOpenConnectionEvent(__p1, __p2);
        }

        #endregion

        #region SetupEventBinding
        
        public void SetupEventBinding(XlApplication application)
        {

            if (true == XlLateBindingApiSettings.EventsEnabled)
            {
                _application = application;
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)application.COMReference;
                Guid guid = new Guid("{00024413-0000-0000-C000-000000000046}");
                connectionPointContainer.FindConnectionPoint(ref guid, out _connectionPoint);
                _connectionPoint.Advise(this, out _connectionCookie);
            }

        }

        public void RemoveEventBinding()
        {
            if (_connectionCookie != 0)
            {
                _connectionPoint.Unadvise(_connectionCookie);
                Marshal.ReleaseComObject(_connectionPoint);
                _connectionPoint = null;
                _connectionCookie = 0;
            }
        }
        
        #endregion
    }
}
