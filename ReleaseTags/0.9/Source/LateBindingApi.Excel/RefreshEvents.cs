using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Excel
{
	#region Event Delegates

	public delegate void RefreshEvents_BeforeRefreshEventHandler(ref bool cancel);
	public delegate void RefreshEvents_AfterRefreshEventHandler(bool success);


	#endregion  
	
	#region InterfaceEvents

	public interface RefreshEvents_Event : IEventBinding 
	{
		event RefreshEvents_BeforeRefreshEventHandler BeforeRefreshEvent;
		event RefreshEvents_AfterRefreshEventHandler AfterRefreshEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	[ComImport, Guid("0002441B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface RefreshEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1596)]
		void BeforeRefresh([In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1597)]
		void AfterRefresh([In] bool success);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class RefreshEvents_SinkHelper : SinkHelper, RefreshEvents
	{
		#region Fields

		private readonly string _riid = "0002441B-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public RefreshEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region RefreshEvents Members
		
		public void BeforeRefresh([In] ref bool cancel)
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[1];
			paramArray.SetValue(cancel,0);
			bool isRecieved = _eventBinding.CallEvent("BeforeRefreshEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void AfterRefresh([In] bool success)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = success;
			bool isRecieved = _eventBinding.CallEvent("AfterRefreshEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		#endregion
	}
    
	#endregion
}

