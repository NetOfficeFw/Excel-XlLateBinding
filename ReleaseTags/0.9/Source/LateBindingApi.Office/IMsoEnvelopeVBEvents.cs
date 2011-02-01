using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Office
{
	#region Event Delegates

	public delegate void IMsoEnvelopeVBEvents_EnvelopeShowEventHandler();
	public delegate void IMsoEnvelopeVBEvents_EnvelopeHideEventHandler();


	#endregion  
	
	#region InterfaceEvents

	public interface IMsoEnvelopeVBEvents_Event : IEventBinding 
	{
		event IMsoEnvelopeVBEvents_EnvelopeShowEventHandler EnvelopeShowEvent;
		event IMsoEnvelopeVBEvents_EnvelopeHideEventHandler EnvelopeHideEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	[ComImport, Guid("000672AD-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface IMsoEnvelopeVBEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void EnvelopeShow();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void EnvelopeHide();

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class IMsoEnvelopeVBEvents_SinkHelper : SinkHelper, IMsoEnvelopeVBEvents
	{
		#region Fields

		private readonly string _riid = "000672AD-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public IMsoEnvelopeVBEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region IMsoEnvelopeVBEvents Members
		
		public void EnvelopeShow()
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("EnvelopeShowEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void EnvelopeHide()
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("EnvelopeHideEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}



		#endregion
	}
    
	#endregion
}

