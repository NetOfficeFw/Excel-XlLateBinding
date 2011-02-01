using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Office
{
	#region Event Delegates

	public delegate void _CustomTaskPaneEvents_VisibleStateChangeEventHandler(LateBindingApi.Office._CustomTaskPane customTaskPaneInst);
	public delegate void _CustomTaskPaneEvents_DockPositionStateChangeEventHandler(LateBindingApi.Office._CustomTaskPane customTaskPaneInst);


	#endregion  
	
	#region InterfaceEvents

	public interface _CustomTaskPaneEvents_Event : IEventBinding 
	{
		event _CustomTaskPaneEvents_VisibleStateChangeEventHandler VisibleStateChangeEvent;
		event _CustomTaskPaneEvents_DockPositionStateChangeEventHandler DockPositionStateChangeEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OF12","OF14")]
	[ComImport, Guid("000C033C-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomTaskPaneEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void VisibleStateChange([In, MarshalAs(UnmanagedType.Interface)] object customTaskPaneInst);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void DockPositionStateChange([In, MarshalAs(UnmanagedType.Interface)] object customTaskPaneInst);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomTaskPaneEvents_SinkHelper : SinkHelper, _CustomTaskPaneEvents
	{
		#region Fields

		private readonly string _riid = "000C033C-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _CustomTaskPaneEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _CustomTaskPaneEvents Members
		
		public void VisibleStateChange([In, MarshalAs(UnmanagedType.Interface)] object customTaskPaneInst)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Office._CustomTaskPane(_eventClass,customTaskPaneInst);
			bool isRecieved = _eventBinding.CallEvent("VisibleStateChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void DockPositionStateChange([In, MarshalAs(UnmanagedType.Interface)] object customTaskPaneInst)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Office._CustomTaskPane(_eventClass,customTaskPaneInst);
			bool isRecieved = _eventBinding.CallEvent("DockPositionStateChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}



		#endregion
	}
    
	#endregion
}

