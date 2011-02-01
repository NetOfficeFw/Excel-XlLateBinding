using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace LateBindingApi.Excel.Charts
{
    #region Event Delegates

    public delegate void ChartEvents_ActivateEventHandler();
    public delegate void ChartEvents_BeforeDoubleClickEventHandler([In] int ElementID, [In] int Arg1, [In] int Arg2, [In] ref bool Cancel);
    public delegate void ChartEvents_BeforeRightClickEventHandler([In] ref bool Cancel);
    public delegate void ChartEvents_CalculateEventHandler();
    public delegate void ChartEvents_DeactivateEventHandler();
    public delegate void ChartEvents_DragOverEventHandler();
    public delegate void ChartEvents_DragPlotEventHandler();
    public delegate void ChartEvents_MouseDownEventHandler([In] int Button, [In] int Shift, [In] int x, [In] int y);
    public delegate void ChartEvents_MouseMoveEventHandler([In] int Button, [In] int Shift, [In] int x, [In] int y);
    public delegate void ChartEvents_MouseUpEventHandler([In] int Button, [In] int Shift, [In] int x, [In] int y);
    public delegate void ChartEvents_ResizeEventHandler();
    public delegate void ChartEvents_SelectEventHandler([In] int ElementID, [In] int Arg1, [In] int Arg2);
    public delegate void ChartEvents_SeriesChangeEventHandler([In] int SeriesIndex, [In] int PointIndex);
    
    #endregion


    [ComImport, Guid("0002440F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface IChartEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x130)]
        void Activate();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fa)]
        void Deactivate();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x100)]
        void Resize();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fb)]
        void MouseDown([In] int Button, [In] int Shift, [In] int x, [In] int y);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fc)]
        void MouseUp([In] int Button, [In] int Shift, [In] int x, [In] int y);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fd)]
        void MouseMove([In] int Button, [In] int Shift, [In] int x, [In] int y);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5fe)]
        void BeforeRightClick([In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x5ff)]
        void DragPlot();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x600)]
        void DragOver();
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x601)]
        void BeforeDoubleClick([In] int ElementID, [In] int Arg1, [In] int Arg2, [In] ref bool Cancel);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0xeb)]
        void Select([In] int ElementID, [In] int Arg1, [In] int Arg2);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x602)]
        void SeriesChange([In] int SeriesIndex, [In] int PointIndex);
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0x117)]
        void Calculate();

    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class XlChartEvents : IChartEvents
    {
        #region Fields

        private XlChart _chart;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion

        #region Construction

        public XlChartEvents()
        {
        }
        
        #endregion

        #region Events

        public void Activate()
        {
            _chart.RaiseActivateEvent();
        }

        public void Deactivate()
        {
            _chart.RaiseDeactivateEvent();
        }

        public void Resize()
        {
            _chart.RaiseResizeEvent();
        }

        public void MouseDown(int Button, int Shift, int x, int y)
        {
            _chart.RaiseMouseDownEvent(Button, Shift,x,y);
        }

        public void MouseUp(int Button, int Shift, int x, int y)
        {
            _chart.RaiseMouseUpEvent(Button, Shift, x, y);
        }

        public void MouseMove(int Button, int Shift, int x, int y)
        {
            _chart.RaiseMouseMoveEvent(Button, Shift, x, y);
        }

        public void BeforeRightClick(ref bool Cancel)
        {
            _chart.RaiseBeforeRightClickEvent(ref Cancel);
        }

        public void DragPlot()
        {
            _chart.RaiseDragPlotEvent();
        }

        public void DragOver()
        {
            _chart.RaiseDragOverEvent();
        }

        public void BeforeDoubleClick(int ElementID, int Arg1, int Arg2, ref bool Cancel)
        {
            _chart.RaiseBeforeDoubleClickEvent(ElementID, Arg1, Arg2, ref Cancel);
        }

        public void Select(int ElementID, int Arg1, int Arg2)
        {
            _chart.RaiseSelectEvent(ElementID, Arg1, Arg2);
        }

        public void SeriesChange(int SeriesIndex, int PointIndex)
        {
            _chart.RaiseSeriesChangeEvent(SeriesIndex, PointIndex);
        }

        public void Calculate()
        {
            _chart.RaiseCalculateEvent();
        }

        #endregion

        #region SetupEventBinding

        public void SetupEventBinding(XlChart chart)
        {

            if (true == XlLateBindingApiSettings.EventsEnabled)
            {
                _chart = chart;
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)chart.COMReference;
                Guid guid = new Guid("{0002440F-0000-0000-C000-000000000046}");
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
