using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.VBIDE
{
	#region Event Delegates

	public delegate void _dispReferencesEvents_ItemAddedEventHandler(LateBindingApi.VBIDE.Reference reference);
	public delegate void _dispReferencesEvents_ItemRemovedEventHandler(LateBindingApi.VBIDE.Reference reference);


	#endregion  
	
	#region InterfaceEvents

	public interface _dispReferencesEvents_Event : IEventBinding 
	{
		event _dispReferencesEvents_ItemAddedEventHandler ItemAddedEvent;
		event _dispReferencesEvents_ItemRemovedEventHandler ItemRemovedEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("VBE")]
	[ComImport, Guid("0002E118-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispReferencesEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void ItemAdded([In, MarshalAs(UnmanagedType.Interface)] object reference);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void ItemRemoved([In, MarshalAs(UnmanagedType.Interface)] object reference);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispReferencesEvents_SinkHelper : SinkHelper, _dispReferencesEvents
	{
		#region Fields

		private readonly string _riid = "0002E118-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _dispReferencesEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _dispReferencesEvents Members
		
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

