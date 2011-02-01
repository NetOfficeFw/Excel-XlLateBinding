using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Excel
{
	#region Event Delegates

	public delegate void WorkbookEvents_OpenEventHandler();
	public delegate void WorkbookEvents_ActivateEventHandler();
	public delegate void WorkbookEvents_DeactivateEventHandler();
	public delegate void WorkbookEvents_BeforeCloseEventHandler(ref bool cancel);
	public delegate void WorkbookEvents_BeforeSaveEventHandler(bool saveAsUI, ref bool cancel);
	public delegate void WorkbookEvents_BeforePrintEventHandler(ref bool cancel);
	public delegate void WorkbookEvents_NewSheetEventHandler(COMObject sh);
	public delegate void WorkbookEvents_AddinInstallEventHandler();
	public delegate void WorkbookEvents_AddinUninstallEventHandler();
	public delegate void WorkbookEvents_WindowResizeEventHandler(LateBindingApi.Excel.Window wn);
	public delegate void WorkbookEvents_WindowActivateEventHandler(LateBindingApi.Excel.Window wn);
	public delegate void WorkbookEvents_WindowDeactivateEventHandler(LateBindingApi.Excel.Window wn);
	public delegate void WorkbookEvents_SheetSelectionChangeEventHandler(COMObject sh, LateBindingApi.Excel.Range target);
	public delegate void WorkbookEvents_SheetBeforeDoubleClickEventHandler(COMObject sh, LateBindingApi.Excel.Range target, ref bool cancel);
	public delegate void WorkbookEvents_SheetBeforeRightClickEventHandler(COMObject sh, LateBindingApi.Excel.Range target, ref bool cancel);
	public delegate void WorkbookEvents_SheetActivateEventHandler(COMObject sh);
	public delegate void WorkbookEvents_SheetDeactivateEventHandler(COMObject sh);
	public delegate void WorkbookEvents_SheetCalculateEventHandler(COMObject sh);
	public delegate void WorkbookEvents_SheetChangeEventHandler(COMObject sh, LateBindingApi.Excel.Range target);
	public delegate void WorkbookEvents_SheetFollowHyperlinkEventHandler(COMObject sh, LateBindingApi.Excel.Hyperlink target);
	public delegate void WorkbookEvents_SheetPivotTableUpdateEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable target);
	public delegate void WorkbookEvents_PivotTableCloseConnectionEventHandler(LateBindingApi.Excel.PivotTable target);
	public delegate void WorkbookEvents_PivotTableOpenConnectionEventHandler(LateBindingApi.Excel.PivotTable target);
	public delegate void WorkbookEvents_SyncEventHandler(LateBindingApi.Office.Enums.MsoSyncEventType syncEventType);
	public delegate void WorkbookEvents_BeforeXmlImportEventHandler(LateBindingApi.Excel.XmlMap map, string url, bool isRefresh, ref bool cancel);
	public delegate void WorkbookEvents_AfterXmlImportEventHandler(LateBindingApi.Excel.XmlMap map, bool isRefresh, LateBindingApi.Excel.Enums.XlXmlImportResult result);
	public delegate void WorkbookEvents_BeforeXmlExportEventHandler(LateBindingApi.Excel.XmlMap map, string url, ref bool cancel);
	public delegate void WorkbookEvents_AfterXmlExportEventHandler(LateBindingApi.Excel.XmlMap map, string url, LateBindingApi.Excel.Enums.XlXmlExportResult result);
	public delegate void WorkbookEvents_RowsetCompleteEventHandler(string description, string sheet, bool success);
	public delegate void WorkbookEvents_SheetPivotTableAfterValueChangeEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, LateBindingApi.Excel.Range targetRange);
	public delegate void WorkbookEvents_SheetPivotTableBeforeAllocateChangesEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void WorkbookEvents_SheetPivotTableBeforeCommitChangesEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void WorkbookEvents_SheetPivotTableBeforeDiscardChangesEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
	public delegate void WorkbookEvents_SheetPivotTableChangeSyncEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable target);
	public delegate void WorkbookEvents_AfterSaveEventHandler(bool success);
	public delegate void WorkbookEvents_NewChartEventHandler(LateBindingApi.Excel.Chart ch);


	#endregion  
	
	#region InterfaceEvents

	public interface WorkbookEvents_Event : IEventBinding 
	{
		event WorkbookEvents_OpenEventHandler OpenEvent;
		event WorkbookEvents_ActivateEventHandler ActivateEvent;
		event WorkbookEvents_DeactivateEventHandler DeactivateEvent;
		event WorkbookEvents_BeforeCloseEventHandler BeforeCloseEvent;
		event WorkbookEvents_BeforeSaveEventHandler BeforeSaveEvent;
		event WorkbookEvents_BeforePrintEventHandler BeforePrintEvent;
		event WorkbookEvents_NewSheetEventHandler NewSheetEvent;
		event WorkbookEvents_AddinInstallEventHandler AddinInstallEvent;
		event WorkbookEvents_AddinUninstallEventHandler AddinUninstallEvent;
		event WorkbookEvents_WindowResizeEventHandler WindowResizeEvent;
		event WorkbookEvents_WindowActivateEventHandler WindowActivateEvent;
		event WorkbookEvents_WindowDeactivateEventHandler WindowDeactivateEvent;
		event WorkbookEvents_SheetSelectionChangeEventHandler SheetSelectionChangeEvent;
		event WorkbookEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent;
		event WorkbookEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent;
		event WorkbookEvents_SheetActivateEventHandler SheetActivateEvent;
		event WorkbookEvents_SheetDeactivateEventHandler SheetDeactivateEvent;
		event WorkbookEvents_SheetCalculateEventHandler SheetCalculateEvent;
		event WorkbookEvents_SheetChangeEventHandler SheetChangeEvent;
		event WorkbookEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent;
		event WorkbookEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent;
		event WorkbookEvents_PivotTableCloseConnectionEventHandler PivotTableCloseConnectionEvent;
		event WorkbookEvents_PivotTableOpenConnectionEventHandler PivotTableOpenConnectionEvent;
		event WorkbookEvents_SyncEventHandler SyncEvent;
		event WorkbookEvents_BeforeXmlImportEventHandler BeforeXmlImportEvent;
		event WorkbookEvents_AfterXmlImportEventHandler AfterXmlImportEvent;
		event WorkbookEvents_BeforeXmlExportEventHandler BeforeXmlExportEvent;
		event WorkbookEvents_AfterXmlExportEventHandler AfterXmlExportEvent;
		event WorkbookEvents_RowsetCompleteEventHandler RowsetCompleteEvent;
		event WorkbookEvents_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent;
		event WorkbookEvents_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent;
		event WorkbookEvents_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent;
		event WorkbookEvents_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent;
		event WorkbookEvents_SheetPivotTableChangeSyncEventHandler SheetPivotTableChangeSyncEvent;
		event WorkbookEvents_AfterSaveEventHandler AfterSaveEvent;
		event WorkbookEvents_NewChartEventHandler NewChartEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	[ComImport, Guid("00024412-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface WorkbookEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(682)]
		void Open();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(304)]
		void Activate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1530)]
		void Deactivate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1546)]
		void BeforeClose([In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1547)]
		void BeforeSave([In] bool saveAsUI, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1549)]
		void BeforePrint([In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1550)]
		void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1552)]
		void AddinInstall();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1553)]
		void AddinUninstall();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void WindowResize([In, MarshalAs(UnmanagedType.Interface)] object wn);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
		void WindowActivate([In, MarshalAs(UnmanagedType.Interface)] object wn);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wn);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1558)]
		void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1559)]
		void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1560)]
		void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1561)]
		void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1562)]
		void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1563)]
		void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1564)]
		void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
		void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2158)]
		void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2159)]
		void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2266)]
		void Sync([In] LateBindingApi.Office.Enums.MsoSyncEventType syncEventType);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2283)]
		void BeforeXmlImport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] bool isRefresh, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2285)]
		void AfterXmlImport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] bool isRefresh, [In] LateBindingApi.Excel.Enums.XlXmlImportResult result);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2287)]
		void BeforeXmlExport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2288)]
		void AfterXmlExport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] LateBindingApi.Excel.Enums.XlXmlExportResult result);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2610)]
		void RowsetComplete([In] string description, [In] string sheet, [In] bool success);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
		void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In, MarshalAs(UnmanagedType.Interface)] object targetRange);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
		void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
		void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
		void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2899)]
		void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2900)]
		void AfterSave([In] bool success);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2901)]
		void NewChart([In, MarshalAs(UnmanagedType.Interface)] object ch);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class WorkbookEvents_SinkHelper : SinkHelper, WorkbookEvents
	{
		#region Fields

		private readonly string _riid = "00024412-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public WorkbookEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region WorkbookEvents Members
		
		public void Open()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("OpenEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void Activate()
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("ActivateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void Deactivate()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("DeactivateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void BeforeClose([In] ref bool cancel)
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[1];
			paramArray.SetValue(cancel,0);
			bool isRecieved = _eventBinding.CallEvent("BeforeCloseEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void BeforeSave([In] bool saveAsUI, [In] ref bool cancel)
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = saveAsUI;
			paramArray.SetValue(cancel,1);
			bool isRecieved = _eventBinding.CallEvent("BeforeSaveEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void BeforePrint([In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[1];
			paramArray.SetValue(cancel,0);
			bool isRecieved = _eventBinding.CallEvent("BeforePrintEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void NewSheet([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new COMObject(_eventClass,sh);
			bool isRecieved = _eventBinding.CallEvent("NewSheetEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void AddinInstall()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("AddinInstallEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void AddinUninstall()
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			bool isRecieved = _eventBinding.CallEvent("AddinUninstallEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
		}

		public void WindowResize([In, MarshalAs(UnmanagedType.Interface)] object wn)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wn);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Window(_eventClass,wn);
			bool isRecieved = _eventBinding.CallEvent("WindowResizeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.Interface)] object wn)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wn);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Window(_eventClass,wn);
			bool isRecieved = _eventBinding.CallEvent("WindowActivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wn)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wn);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Window(_eventClass,wn);
			bool isRecieved = _eventBinding.CallEvent("WindowDeactivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetSelectionChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.Range(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("SheetSelectionChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetBeforeDoubleClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.Range(_eventClass,target);
			paramArray.SetValue(cancel,2);
			bool isRecieved = _eventBinding.CallEvent("SheetBeforeDoubleClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetBeforeRightClick([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.Range(_eventClass,target);
			paramArray.SetValue(cancel,2);
			bool isRecieved = _eventBinding.CallEvent("SheetBeforeRightClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetActivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new COMObject(_eventClass,sh);
			bool isRecieved = _eventBinding.CallEvent("SheetActivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetDeactivate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new COMObject(_eventClass,sh);
			bool isRecieved = _eventBinding.CallEvent("SheetDeactivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetCalculate([In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new COMObject(_eventClass,sh);
			bool isRecieved = _eventBinding.CallEvent("SheetCalculateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.Range(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("SheetChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.Hyperlink(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("SheetFollowHyperlinkEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("SheetPivotTableUpdateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableCloseConnection([In, MarshalAs(UnmanagedType.Interface)] object target)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("PivotTableCloseConnectionEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void PivotTableOpenConnection([In, MarshalAs(UnmanagedType.Interface)] object target)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("PivotTableOpenConnectionEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void Sync([In] LateBindingApi.Office.Enums.MsoSyncEventType syncEventType)
		{
			object[] paramArray = new object[1];
			paramArray[0] = syncEventType;
			bool isRecieved = _eventBinding.CallEvent("SyncEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void BeforeXmlImport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] bool isRefresh, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[1] = url;
			paramArray[2] = isRefresh;
			paramArray.SetValue(cancel,3);
			bool isRecieved = _eventBinding.CallEvent("BeforeXmlImportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void AfterXmlImport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] bool isRefresh, [In] LateBindingApi.Excel.Enums.XlXmlImportResult result)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[1] = isRefresh;
			paramArray[2] = result;
			bool isRecieved = _eventBinding.CallEvent("AfterXmlImportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void BeforeXmlExport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] ref bool cancel)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[1] = url;
			paramArray.SetValue(cancel,2);
			bool isRecieved = _eventBinding.CallEvent("BeforeXmlExportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void AfterXmlExport([In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] LateBindingApi.Excel.Enums.XlXmlExportResult result)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[1] = url;
			paramArray[2] = result;
			bool isRecieved = _eventBinding.CallEvent("AfterXmlExportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void RowsetComplete([In] string description, [In] string sheet, [In] bool success)
		{
			object[] paramArray = new object[3];
			paramArray[0] = description;
			paramArray[1] = sheet;
			paramArray[2] = success;
			bool isRecieved = _eventBinding.CallEvent("RowsetCompleteEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In, MarshalAs(UnmanagedType.Interface)] object targetRange)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(targetPivotTable);
                Marshal.ReleaseComObject(targetRange);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[2] = new LateBindingApi.Excel.Range(_eventClass,targetRange);
			bool isRecieved = _eventBinding.CallEvent("SheetPivotTableAfterValueChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(targetPivotTable);
                return;
            }

			object[] paramArray = new object[5];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[2] = valueChangeStart;
			paramArray[3] = valueChangeEnd;
			paramArray.SetValue(cancel,4);
			bool isRecieved = _eventBinding.CallEvent("SheetPivotTableBeforeAllocateChangesEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(targetPivotTable);
                return;
            }

			object[] paramArray = new object[5];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[2] = valueChangeStart;
			paramArray[3] = valueChangeEnd;
			paramArray.SetValue(cancel,4);
			bool isRecieved = _eventBinding.CallEvent("SheetPivotTableBeforeCommitChangesEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(targetPivotTable);
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,targetPivotTable);
			paramArray[2] = valueChangeStart;
			paramArray[3] = valueChangeEnd;
			bool isRecieved = _eventBinding.CallEvent("SheetPivotTableBeforeDiscardChangesEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void SheetPivotTableChangeSync([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(sh);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new COMObject(_eventClass,sh);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("SheetPivotTableChangeSyncEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void AfterSave([In] bool success)
        {
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = success;
			bool isRecieved = _eventBinding.CallEvent("AfterSaveEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void NewChart([In, MarshalAs(UnmanagedType.Interface)] object ch)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(ch);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Chart(_eventClass,ch);
			bool isRecieved = _eventBinding.CallEvent("NewChartEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		#endregion
	}
    
	#endregion
}

