using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.VBIDE
{
	#region Event Delegates

	public delegate void _dispReferences_Events_ItemAddedEventHandler(LateBindingApi.VBIDE.Reference reference);
	public delegate void _dispReferences_Events_ItemRemovedEventHandler(LateBindingApi.VBIDE.Reference reference);


	#endregion  
	
	#region InterfaceEvents

	public interface _dispReferences_Events_Event : IEventBinding 
	{
		event _dispReferences_Events_ItemAddedEventHandler ItemAddedEvent;
		event _dispReferences_Events_ItemRemovedEventHandler ItemRemovedEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("VBE")]
	[ComImport, Guid("CDDE3804-2064-11CF-867F-00AA005FF34A"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferences_Events
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(0)]
		void ItemAdded([In, MarshalAs(UnmanagedType.Interface)] object reference);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.Interface)] object reference);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispReferences_Events_SinkHelper : SinkHelper, _dispReferences_Events
	{
		#region Fields

		private readonly string _riid = "CDDE3804-2064-11CF-867F-00AA005FF34A";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _dispReferences_Events_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _dispReferences_Events Members
		
		public void ItemAdded([In, MarshalAs(UnmanagedType.Interface)] object reference)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(reference);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.VBIDE.Reference(_eventClass,reference);
			bool isRecieved = _eventBinding.CallEvent("ItemAddedEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void ItemRemoved([In, MarshalAs(UnmanagedType.Interface)] object reference)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(reference);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.VBIDE.Reference(_eventClass,reference);
			bool isRecieved = _eventBinding.CallEvent("ItemRemovedEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		#endregion
	}
    
	#endregion
}

