using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Excel
{
	#region Event Delegates

	public delegate void ChartEvents_ActivateEventHandler();
	public delegate void ChartEvents_DeactivateEventHandler();
	public delegate void ChartEvents_ResizeEventHandler();
	public delegate void ChartEvents_MouseDownEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void ChartEvents_MouseUpEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void ChartEvents_MouseMoveEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void ChartEvents_BeforeRightClickEventHandler(ref bool cancel);
	public delegate void ChartEvents_DragPlotEventHandler();
	public delegate void ChartEvents_DragOverEventHandler();
	public delegate void ChartEvents_BeforeDoubleClickEventHandler(Int32 elementID, Int32 arg1, Int32 arg2, ref bool cancel);
	public delegate void ChartEvents_SelectEventHandler(Int32 elementID, Int32 arg1, Int32 arg2);
	public delegate void ChartEvents_SeriesChangeEventHandler(Int32 seriesIndex, Int32 pointIndex);
	public delegate void ChartEvents_CalculateEventHandler();


	#endregion  
	
	#region InterfaceEvents

	public interface ChartEvents_Event : IEventBinding 
	{
		event ChartEvents_ActivateEventHandler ActivateEvent;
		event ChartEvents_DeactivateEventHandler DeactivateEvent;
		event ChartEvents_ResizeEventHandler ResizeEvent;
		event ChartEvents_MouseDownEventHandler MouseDownEvent;
		event ChartEvents_MouseUpEventHandler MouseUpEvent;
		event ChartEvents_MouseMoveEventHandler MouseMoveEvent;
		event ChartEvents_BeforeRightClickEventHandler BeforeRightClickEvent;
		event ChartEvents_DragPlotEventHandler DragPlotEvent;
		event ChartEvents_DragOverEventHandler DragOverEvent;
		event ChartEvents_BeforeDoubleClickEventHandler BeforeDoubleClickEvent;
		event ChartEvents_SelectEventHandler SelectEvent;
		event ChartEvents_SeriesChangeEventHandler SeriesChangeEvent;
		event ChartEvents_CalculateEventHandler CalculateEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	[ComImport, Guid("0002440F-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface ChartEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(256)]
		void Resize();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1531)]
		void MouseDown([In] Int32 button, [In] Int32 shift, [In] Int32 x, [In] Int32 y);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1532)]
		void MouseUp([In] Int32 button, [In] Int32 shift, [In] Int32 x, [In] Int32 y);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1533)]
		void MouseMove([In] Int32 button, [In] Int32 shift, [In] Int32 x, [In] Int32 y);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
		void BeforeRightClick([In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1535)]
		void DragPlot();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1536)]
		void DragOver();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
		void BeforeDoubleClick([In] Int32 elementID, [In] Int32 arg1, [In] Int32 arg2, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(235)]
		void Select([In] Int32 elementID, [In] Int32 arg1, [In] Int32 arg2);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1538)]
		void SeriesChange([In] Int32 seriesIndex, [In] Int32 pointIndex);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
		void Calculate();

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class ChartEvents_SinkHelper : SinkHelper, ChartEvents
	{
		#region Fields

		private readonly string _riid = "0002440F-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public ChartEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region ChartEvents Members
		
		public void Activate()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("ActivateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void Deactivate()
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("DeactivateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void Resize()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("ResizeEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void MouseDown([In] Int32 button, [In] Int32 shift, [In] Int32 x, [In] Int32 y)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = button;
			paramArray[1] = shift;
			paramArray[2] = x;
			paramArray[3] = y;
			bool isRecieved = _eventBinding.CallEvent("MouseDownEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void MouseUp([In] Int32 button, [In] Int32 shift, [In] Int32 x, [In] Int32 y)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = button;
			paramArray[1] = shift;
			paramArray[2] = x;
			paramArray[3] = y;
			bool isRecieved = _eventBinding.CallEvent("MouseUpEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void MouseMove([In] Int32 button, [In] Int32 shift, [In] Int32 x, [In] Int32 y)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = button;
			paramArray[1] = shift;
			paramArray[2] = x;
			paramArray[3] = y;
			bool isRecieved = _eventBinding.CallEvent("MouseMoveEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void BeforeRightClick([In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[1];
			paramArray.SetValue(cancel,0);
			bool isRecieved = _eventBinding.CallEvent("BeforeRightClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void DragPlot()
		{
			bool isRecieved = _eventBinding.CallEvent("DragPlotEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void DragOver()
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("DragOverEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void BeforeDoubleClick([In] Int32 elementID, [In] Int32 arg1, [In] Int32 arg2, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = elementID;
			paramArray[1] = arg1;
			paramArray[2] = arg2;
			paramArray.SetValue(cancel,3);
			bool isRecieved = _eventBinding.CallEvent("BeforeDoubleClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void Select([In] Int32 elementID, [In] Int32 arg1, [In] Int32 arg2)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = elementID;
			paramArray[1] = arg1;
			paramArray[2] = arg2;
			bool isRecieved = _eventBinding.CallEvent("SelectEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SeriesChange([In] Int32 seriesIndex, [In] Int32 pointIndex)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = seriesIndex;
			paramArray[1] = pointIndex;
			bool isRecieved = _eventBinding.CallEvent("SeriesChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void Calculate()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("CalculateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		#endregion
	}
    
	#endregion
}

