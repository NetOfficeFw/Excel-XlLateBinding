using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Office
{
	#region Event Delegates

	public delegate void _CustomXMLPartEvents_NodeAfterInsertEventHandler(LateBindingApi.Office.CustomXMLNode newNode, bool inUndoRedo);
	public delegate void _CustomXMLPartEvents_NodeAfterDeleteEventHandler(LateBindingApi.Office.CustomXMLNode oldNode, LateBindingApi.Office.CustomXMLNode oldParentNode, LateBindingApi.Office.CustomXMLNode oldNextSibling, bool inUndoRedo);
	public delegate void _CustomXMLPartEvents_NodeAfterReplaceEventHandler(LateBindingApi.Office.CustomXMLNode oldNode, LateBindingApi.Office.CustomXMLNode newNode, bool inUndoRedo);


	#endregion  
	
	#region InterfaceEvents

	public interface _CustomXMLPartEvents_Event : IEventBinding 
	{
		event _CustomXMLPartEvents_NodeAfterInsertEventHandler NodeAfterInsertEvent;
		event _CustomXMLPartEvents_NodeAfterDeleteEventHandler NodeAfterDeleteEvent;
		event _CustomXMLPartEvents_NodeAfterReplaceEventHandler NodeAfterReplaceEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OF12","OF14")]
	[ComImport, Guid("000CDB07-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CustomXMLPartEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void NodeAfterInsert([In, MarshalAs(UnmanagedType.Interface)] object newNode, [In] bool inUndoRedo);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void NodeAfterDelete([In, MarshalAs(UnmanagedType.Interface)] object oldNode, [In, MarshalAs(UnmanagedType.Interface)] object oldParentNode, [In, MarshalAs(UnmanagedType.Interface)] object oldNextSibling, [In] bool inUndoRedo);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void NodeAfterReplace([In, MarshalAs(UnmanagedType.Interface)] object oldNode, [In, MarshalAs(UnmanagedType.Interface)] object newNode, [In] bool inUndoRedo);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CustomXMLPartEvents_SinkHelper : SinkHelper, _CustomXMLPartEvents
	{
		#region Fields

		private readonly string _riid = "000CDB07-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _CustomXMLPartEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _CustomXMLPartEvents Members
		
		public void NodeAfterInsert([In, MarshalAs(UnmanagedType.Interface)] object newNode, [In] bool inUndoRedo)
		{
			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Office.CustomXMLNode(_eventClass,newNode);
			paramArray[1] = inUndoRedo;
			bool isRecieved = _eventBinding.CallEvent("NodeAfterInsertEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void NodeAfterDelete([In, MarshalAs(UnmanagedType.Interface)] object oldNode, [In, MarshalAs(UnmanagedType.Interface)] object oldParentNode, [In, MarshalAs(UnmanagedType.Interface)] object oldNextSibling, [In] bool inUndoRedo)
		{
			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Office.CustomXMLNode(_eventClass,oldNode);
			paramArray[1] = new LateBindingApi.Office.CustomXMLNode(_eventClass,oldParentNode);
			paramArray[2] = new LateBindingApi.Office.CustomXMLNode(_eventClass,oldNextSibling);
			paramArray[3] = inUndoRedo;
			bool isRecieved = _eventBinding.CallEvent("NodeAfterDeleteEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void NodeAfterReplace([In, MarshalAs(UnmanagedType.Interface)] object oldNode, [In, MarshalAs(UnmanagedType.Interface)] object newNode, [In] bool inUndoRedo)
		{
			object[] paramArray = new object[3];
			paramArray[0] = new LateBindingApi.Office.CustomXMLNode(_eventClass,oldNode);
			paramArray[1] = new LateBindingApi.Office.CustomXMLNode(_eventClass,newNode);
			paramArray[2] = inUndoRedo;
			bool isRecieved = _eventBinding.CallEvent("NodeAfterReplaceEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}



		#endregion
	}
    
	#endregion
}

