using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Excel
{
	#region Event Delegates

	public delegate void OLEObjectEvents_GotFocusEventHandler();
	public delegate void OLEObjectEvents_LostFocusEventHandler();


	#endregion  
	
	#region InterfaceEvents

	public interface OLEObjectEvents_Event : IEventBinding 
	{
		event OLEObjectEvents_GotFocusEventHandler GotFocusEvent;
		event OLEObjectEvents_LostFocusEventHandler LostFocusEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	[ComImport, Guid("00024410-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface OLEObjectEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1541)]
		void GotFocus();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1542)]
		void LostFocus();

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class OLEObjectEvents_SinkHelper : SinkHelper, OLEObjectEvents
	{
		#region Fields

		private readonly string _riid = "00024410-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public OLEObjectEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region OLEObjectEvents Members
		
		public void GotFocus()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("GotFocusEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void LostFocus()
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("LostFocusEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		#endregion
	}
    
	#endregion
}

