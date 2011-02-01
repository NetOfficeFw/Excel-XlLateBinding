using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Office
{
	#region Event Delegates

	public delegate void _CustomXMLPartsEvents_PartAfterAddEventHandler(LateBindingApi.Office.CustomXMLPart newPart);
	public delegate void _CustomXMLPartsEvents_PartBeforeDeleteEventHandler(LateBindingApi.Office.CustomXMLPart oldPart);
	public delegate void _CustomXMLPartsEvents_PartAfterLoadEventHandler(LateBindingApi.Office.CustomXMLPart part);


	#endregion  
	
	#region InterfaceEvents

	public interface _CustomXMLPartsEvents_Event : IEventBinding 
	{
		event _CustomXMLPartsEvents_PartAfterAddEventHandler PartAfterAddEvent;
		event _CustomXMLPartsEvents_PartBeforeDeleteEventHandler PartBeforeDeleteEvent;
		event _CustomXMLPartsEvents_PartAfterLoadEventHandler PartAfterLoadEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OF12","OF14")]
	[ComImport, Guid("000CDB0B-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartsEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void PartAfterAdd([In, MarshalAs(UnmanagedType.Interface)] object newPart);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void PartBeforeDelete([In, MarshalAs(UnmanagedType.Interface)] object oldPart);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void PartAfterLoad([In, MarshalAs(UnmanagedType.Interface)] object part);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomXMLPartsEvents_SinkHelper : SinkHelper, _CustomXMLPartsEvents
	{
		#region Fields

		private readonly string _riid = "000CDB0B-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _CustomXMLPartsEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _CustomXMLPartsEvents Members
		
		public void PartAfterAdd([In, MarshalAs(UnmanagedType.Interface)] object newPart)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Office.CustomXMLPart(_eventClass,newPart);
			bool isRecieved = _eventBinding.CallEvent("PartAfterAddEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PartBeforeDelete([In, MarshalAs(UnmanagedType.Interface)] object oldPart)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Office.CustomXMLPart(_eventClass,oldPart);
			bool isRecieved = _eventBinding.CallEvent("PartBeforeDeleteEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PartAfterLoad([In, MarshalAs(UnmanagedType.Interface)] object part)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Office.CustomXMLPart(_eventClass,part);
			bool isRecieved = _eventBinding.CallEvent("PartAfterLoadEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}



		#endregion
	}
    
	#endregion
}

