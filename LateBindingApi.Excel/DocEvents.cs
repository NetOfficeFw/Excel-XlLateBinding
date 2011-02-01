using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Excel
{
	#region Event Delegates

	public delegate void DocEvents_SelectionChangeEventHandler(LateBindingApi.Excel.Range target);
	public delegate void DocEvents_BeforeDoubleClickEventHandler(LateBindingApi.Excel.Range target, ref bool cancel);
	public delegate void DocEvents_BeforeRightClickEventHandler(LateBindingApi.Excel.Range target, ref bool cancel);
	public delegate void DocEvents_ActivateEventHandler();
	public delegate void DocEvents_DeactivateEventHandler();
	public delegate void DocEvents_CalculateEventHandler();
	public delegate void DocEvents_ChangeEventHandler(LateBindingApi.Excel.Range target);
	public delegate void DocEvents_FollowHyperlinkEventHandler(LateBindingApi.Excel.Hyperlink target);
	public delegate void DocEvents_PivotTableUpdateEventHandler(LateBindingApi.Excel.PivotTable target);
	public delegate void DocEvents_PivotTableAfterValueChangeEventHandler(LateBindingApi.Excel.PivotTable targetPivotTable, LateBindingApi.Excel.Range targetRange);
	public delegate void DocEvents_PivotTableBeforeAllocateChangesEventHandler(LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void DocEvents_PivotTableBeforeCommitChangesEventHandler(LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void DocEvents_PivotTableBeforeDiscardChangesEventHandler(LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
	public delegate void DocEvents_PivotTableChangeSyncEventHandler(LateBindingApi.Excel.PivotTable target);


	#endregion  
	
	#region InterfaceEvents

	public interface DocEvents_Event : IEventBinding 
	{
		event DocEvents_SelectionChangeEventHandler SelectionChangeEvent;
		event DocEvents_BeforeDoubleClickEventHandler BeforeDoubleClickEvent;
		event DocEvents_BeforeRightClickEventHandler BeforeRightClickEvent;
		event DocEvents_ActivateEventHandler ActivateEvent;
		event DocEvents_DeactivateEventHandler DeactivateEvent;
		event DocEvents_CalculateEventHandler CalculateEvent;
		event DocEvents_ChangeEventHandler ChangeEvent;
		event DocEvents_FollowHyperlinkEventHandler FollowHyperlinkEvent;
		event DocEvents_PivotTableUpdateEventHandler PivotTableUpdateEvent;
		event DocEvents_PivotTableAfterValueChangeEventHandler PivotTableAfterValueChangeEvent;
		event DocEvents_PivotTableBeforeAllocateChangesEventHandler PivotTableBeforeAllocateChangesEvent;
		event DocEvents_PivotTableBeforeCommitChangesEventHandler PivotTableBeforeCommitChangesEvent;
		event DocEvents_PivotTableBeforeDiscardChangesEventHandler PivotTableBeforeDiscardChangesEvent;
		event DocEvents_PivotTableChangeSyncEventHandler PivotTableChangeSyncEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	[ComImport, Guid("00024411-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface DocEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1543)]
		void SelectionChange([In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1537)]
		void BeforeDoubleClick([In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1534)]
		void BeforeRightClick([In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(279)]
		void Calculate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1545)]
		void Change([In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1470)]
		void FollowHyperlink([In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2156)]
		void PivotTableUpdate([In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2886)]
		void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In, MarshalAs(UnmanagedType.Interface)] object targetRange);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2889)]
		void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2892)]
		void PivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2893)]
		void PivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2894)]
		void PivotTableChangeSync([In, MarshalAs(UnmanagedType.Interface)] object target);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class DocEvents_SinkHelper : SinkHelper, DocEvents
	{
		#region Fields

		private readonly string _riid = "00024411-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public DocEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region DocEvents Members
		
		public void SelectionChange([In, MarshalAs(UnmanagedType.Interface)] object target)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Range(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("SelectionChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void BeforeDoubleClick([In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel)
		{
			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Range(_eventClass,target);
			paramArray.SetValue(cancel,1);
			bool isRecieved = _eventBinding.CallEvent("BeforeDoubleClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void BeforeRightClick([In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel)
		{
			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Range(_eventClass,target);
			paramArray.SetValue(cancel,1);
			bool isRecieved = _eventBinding.CallEvent("BeforeRightClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void Activate()
		{
			bool isRecieved = _eventBinding.CallEvent("ActivateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void Deactivate()
		{
			bool isRecieved = _eventBinding.CallEvent("DeactivateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void Calculate()
		{
			bool isRecieved = _eventBinding.CallEvent("CalculateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void Change([In, MarshalAs(UnmanagedType.Interface)] object target)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Range(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("ChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void FollowHyperlink([In, MarshalAs(UnmanagedType.Interface)] object target)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Hyperlink(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("FollowHyperlinkEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableUpdate([In, MarshalAs(UnmanagedType.Interface)] object target)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("PivotTableUpdateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableAfterValueChange([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In, MarshalAs(UnmanagedType.Interface)] object targetRange)
		{
			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[1] = new LateBindingApi.Excel.Range(_eventClass,targetRange);
			bool isRecieved = _eventBinding.CallEvent("PivotTableAfterValueChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel)
		{
			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[1] = valueChangeStart;
			paramArray[2] = valueChangeEnd;
			paramArray.SetValue(cancel,3);
			bool isRecieved = _eventBinding.CallEvent("PivotTableBeforeAllocateChangesEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel)
		{
			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[1] = valueChangeStart;
			paramArray[2] = valueChangeEnd;
			paramArray.SetValue(cancel,3);
			bool isRecieved = _eventBinding.CallEvent("PivotTableBeforeCommitChangesEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd)
		{
			object[] paramArray = new object[3];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[1] = valueChangeStart;
			paramArray[2] = valueChangeEnd;
			bool isRecieved = _eventBinding.CallEvent("PivotTableBeforeDiscardChangesEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableChangeSync([In, MarshalAs(UnmanagedType.Interface)] object target)
		{
			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("PivotTableChangeSyncEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}



		#endregion
	}
    
	#endregion
}

