using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Pivot;
using LateBindingApi.Excel.Windows;

namespace LateBindingApi.Excel
{
    #region EventDelegates

    public delegate void WorkbookEvents_ActivateEventHandler();
    public delegate void WorkbookEvents_AddinInstallEventHandler();
    public delegate void WorkbookEvents_AddinUninstallEventHandler();
    public delegate void WorkbookEvents_BeforeCloseEventHandler([In] ref bool Cancel);
    public delegate void WorkbookEvents_BeforePrintEventHandler([In] ref bool Cancel);
    public delegate void WorkbookEvents_BeforeSaveEventHandler([In] bool SaveAsUI, [In] ref bool Cancel);
    public delegate void WorkbookEvents_DeactivateEventHandler();
    public delegate void WorkbookEvents_NewSheetEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
    public delegate void WorkbookEvents_OpenEventHandler();
    public delegate void WorkbookEvents_PivotTableCloseConnectionEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlPivotTable Target);
    public delegate void WorkbookEvents_PivotTableOpenConnectionEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlPivotTable Target);
    public delegate void WorkbookEvents_SheetActivateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
    public delegate void WorkbookEvents_SheetBeforeDoubleClickEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target, [In] ref bool Cancel);
    public delegate void WorkbookEvents_SheetBeforeRightClickEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target, [In] ref bool Cancel);
    public delegate void WorkbookEvents_SheetCalculateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
    public delegate void WorkbookEvents_SheetChangeEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target);
    public delegate void WorkbookEvents_SheetDeactivateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
    public delegate void WorkbookEvents_SheetFollowHyperlinkEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlHyperlink Target);
    public delegate void WorkbookEvents_SheetPivotTableUpdateEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlPivotTable Target);
    public delegate void WorkbookEvents_SheetSelectionChangeEventHandler([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] XlRange Target);
    public delegate void WorkbookEvents_WindowActivateEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWindow Wn);
    public delegate void WorkbookEvents_WindowDeactivateEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWindow Wn);
    public delegate void WorkbookEvents_WindowResizeEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlWindow Wn);

    #endregion

    [ComImport, Guid("00024412-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface IWorkbookEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x783)]
        void Open();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x130)]
        void Activate();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fa)]
        void Deactivate();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x60a)]
        void BeforeClose([In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x60b)]
        void BeforeSave([In] bool SaveAsUI, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x60d)]
        void BeforePrint([In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x60e)]
        void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x610)]
        void AddinInstall();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x611)]
        void AddinUninstall();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x612)]
        void WindowResize([In, MarshalAs(UnmanagedType.Interface)] object Wn);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x614)]
        void WindowActivate([In, MarshalAs(UnmanagedType.Interface)] object Wn);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x615)]
        void WindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object Wn);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x616)]
        void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x617)]
        void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object Target, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x618)]
        void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object Target, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x619)]
        void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61a)]
        void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61b)]
        void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object Sh);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x61c)]
        void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x73e)]
        void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x86d)]
        void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object Sh, [In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x86e)]
        void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x86f)]
        void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.Interface)] object Target);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class XlWorkbookEvents : IWorkbookEvents
    {
        #region Fields

        private XlWorkbook _workBook;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion
        
        #region Construction

        public XlWorkbookEvents()
        {
        }
        
        #endregion

        #region Events

        public void Open()
        {
            _workBook.RaiseOpenEvent();
        }

        public void Activate()
        {
            _workBook.RaiseActivateEvent();
        }

        public void Deactivate()
        {
            _workBook.RaiseDeactivateEvent();
        }

        public void BeforeClose(ref bool Cancel)
        {
            _workBook.RaiseBeforeCloseEvent(ref Cancel);
        }

        public void BeforeSave(bool SaveAsUI, ref bool Cancel)
        {
            _workBook.RaiseBeforeSaveEvent(SaveAsUI,ref Cancel);
        }

        public void BeforePrint(ref bool Cancel)
        {
            _workBook.RaiseBeforePrintEvent(ref Cancel);
        }

        public void NewSheet(object Sh)
        {
            _workBook.RaiseNewSheetEvent(Sh);
        }

        public void AddinInstall()
        {
            _workBook.RaiseAddinInstallEvent();
        }

        public void AddinUninstall()
        {
            _workBook.RaiseAddinUninstallEvent();
        }

        public void WindowResize(object Wn)
        {
            _workBook.RaiseWindowResizeEvent(Wn);
        }

        public void WindowActivate(object Wn)
        {
            _workBook.RaiseWindowActivateEvent(Wn);
        }

        public void WindowDeactivate(object Wn)
        {
            _workBook.RaiseWindowDeactivateEvent(Wn);
        }

        public void SheetSelectionChange(object Sh, object Target)
        {
            _workBook.RaiseSheetSelectionChangeEvent(Sh, Target);
        }

        public void SheetBeforeDoubleClick(object Sh, object Target, ref bool Cancel)
        {
            _workBook.RaiseSheetBeforeDoubleClickEvent(Sh, Target, ref Cancel);
        }

        public void SheetBeforeRightClick(object Sh, object Target, ref bool Cancel)
        {
            _workBook.RaiseSheetBeforeRightClickEvent(Sh, Target, ref Cancel);
        }

        public void SheetActivate(object Sh)
        {
            _workBook.RaiseSheetActivateEvent(Sh);
        }

        public void SheetDeactivate(object Sh)
        {
            _workBook.RaiseSheetDeactivateEvent(Sh);
        }

        public void SheetCalculate(object Sh)
        {
            _workBook.RaiseSheetCalculateEvent(Sh);
        }

        public void SheetChange(object Sh, object Target)
        {
            _workBook.RaiseSheetChangeEvent(Sh, Target);
        }

        public void SheetFollowHyperlink(object Sh, object Target)
        {
            _workBook.RaiseSheetFollowHyperlinkEvent(Sh, Target);
        }

        public void SheetPivotTableUpdate(object Sh, object Target)
        {
            _workBook.RaiseSheetPivotTableUpdateEvent(Sh, Target);
        }

        public void PivotTableCloseConnection(object Target)
        {
            _workBook.RaisePivotTableCloseConnectionEvent(Target);
        }

        public void PivotTableOpenConnection(object Target)
        {
            _workBook.RaisePivotTableOpenConnectionEvent(Target);
        }

        #endregion

        #region SetupEventBinding

        public void SetupEventBinding(XlWorkbook workBook)
        {

            if (true == XlLateBindingApiSettings.EventsEnabled)
            {
                _workBook = workBook;
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)workBook.COMReference;
                Guid guid = new Guid("{00024412-0000-0000-C000-000000000046}");
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
