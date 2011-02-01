using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Pivot;

namespace LateBindingApi.Excel
{
    #region EventDelegates

    public delegate void DocEvents_ActivateEventHandler();
    public delegate void DocEvents_BeforeDoubleClickEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlRange Target, [In] ref bool Cancel);
    public delegate void DocEvents_BeforeRightClickEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlRange Target, [In] ref bool Cancel);
    public delegate void DocEvents_CalculateEventHandler();
    public delegate void DocEvents_ChangeEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlRange Target);
    public delegate void DocEvents_DeactivateEventHandler();
    public delegate void DocEvents_FollowHyperlinkEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlHyperlink Target);
    public delegate void DocEvents_PivotTableUpdateEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlPivotTable Target);
    public delegate void DocEvents_SelectionChangeEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlRange Target);
    
    #endregion

    [ComImport, Guid("00024411-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface IDocEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x607)]
        void SelectionChange([In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x601)]
        void BeforeDoubleClick([In, MarshalAs(UnmanagedType.Interface)] object Target, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fe)]
        void BeforeRightClick([In, MarshalAs(UnmanagedType.Interface)] object Target, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x130)]
        void Activate();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fa)]
        void Deactivate();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x117)]
        void Calculate();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x609)]
        void Change([In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5be)]
        void FollowHyperlink([In, MarshalAs(UnmanagedType.Interface)] object Target);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x86c)]
        void PivotTableUpdate([In, MarshalAs(UnmanagedType.Interface)] object Target);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class XlWorksheetEvents : IDocEvents
    {
        #region Fields

        private XlWorksheet _workSheet;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion
        
        #region Construction

        public XlWorksheetEvents()
        {
        }
        
        #endregion
   
        #region SetupEventBinding

        public void SetupEventBinding(XlWorksheet workSheet)
        {

            if (true == XlLateBindingApiSettings.EventsEnabled)
            {
                _workSheet = workSheet;
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)workSheet.COMReference;
                Guid guid = new Guid("{00024411-0000-0000-C000-000000000046}");
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

        #region IDocEvents Members

        public void SelectionChange(object Target)
        {
            _workSheet.RaiseSelectionChangeEvent(Target);
        }

        public void BeforeDoubleClick(object Target, ref bool Cancel)
        {
            _workSheet.RaiseBeforeDoubleClickEvent(Target, ref Cancel);
        }

        public void BeforeRightClick(object Target, ref bool Cancel)
        {
            _workSheet.RaiseBeforeRightClickEvent(Target, ref Cancel);
        }

        public void Activate()
        {
            _workSheet.RaiseActivateEvent();
        }

        public void Deactivate()
        {
            _workSheet.RaiseDeactivateEvent();
        }

        public void Calculate()
        {
            _workSheet.RaiseCalculateEvent();
        }

        public void Change(object Target)
        {
            _workSheet.RaiseChangeEvent(Target);
        }

        public void FollowHyperlink(object Target)
        {
            _workSheet.RaiseFollowHyperlinkEvent(Target);
        }

        public void PivotTableUpdate(object Target)
        {
            _workSheet.RaisePivotTableUpdateEvent(Target);

        }

        #endregion
    }
}
