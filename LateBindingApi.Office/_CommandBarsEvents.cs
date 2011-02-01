using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Office
{
	#region Event Delegates

	public delegate void _CommandBarsEvents_OnUpdateEventHandler();


	#endregion  
	
	#region InterfaceEvents

	public interface _CommandBarsEvents_Event : IEventBinding 
	{
		event _CommandBarsEvents_OnUpdateEventHandler OnUpdateEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	[ComImport, Guid("000C0352-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarsEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void OnUpdate();

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CommandBarsEvents_SinkHelper : SinkHelper, _CommandBarsEvents
	{
		#region Fields

		private readonly string _riid = "000C0352-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _CommandBarsEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _CommandBarsEvents Members
		
		public void OnUpdate()
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("OnUpdateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}



		#endregion
	}
    
	#endregion
}

