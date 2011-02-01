using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Excel
{
	#region Event Delegates

	public delegate void AppEvents_NewWorkbookEventHandler(LateBindingApi.Excel.Workbook wb);
	public delegate void AppEvents_SheetSelectionChangeEventHandler(COMObject sh, LateBindingApi.Excel.Range target);
	public delegate void AppEvents_SheetBeforeDoubleClickEventHandler(COMObject sh, LateBindingApi.Excel.Range target, ref bool cancel);
	public delegate void AppEvents_SheetBeforeRightClickEventHandler(COMObject sh, LateBindingApi.Excel.Range target, ref bool cancel);
	public delegate void AppEvents_SheetActivateEventHandler(COMObject sh);
	public delegate void AppEvents_SheetDeactivateEventHandler(COMObject sh);
	public delegate void AppEvents_SheetCalculateEventHandler(COMObject sh);
	public delegate void AppEvents_SheetChangeEventHandler(COMObject sh, LateBindingApi.Excel.Range target);
	public delegate void AppEvents_WorkbookOpenEventHandler(LateBindingApi.Excel.Workbook wb);
	public delegate void AppEvents_WorkbookActivateEventHandler(LateBindingApi.Excel.Workbook wb);
	public delegate void AppEvents_WorkbookDeactivateEventHandler(LateBindingApi.Excel.Workbook wb);
	public delegate void AppEvents_WorkbookBeforeCloseEventHandler(LateBindingApi.Excel.Workbook wb, ref bool cancel);
	public delegate void AppEvents_WorkbookBeforeSaveEventHandler(LateBindingApi.Excel.Workbook wb, bool saveAsUI, ref bool cancel);
	public delegate void AppEvents_WorkbookBeforePrintEventHandler(LateBindingApi.Excel.Workbook wb, ref bool cancel);
	public delegate void AppEvents_WorkbookNewSheetEventHandler(LateBindingApi.Excel.Workbook wb, COMObject sh);
	public delegate void AppEvents_WorkbookAddinInstallEventHandler(LateBindingApi.Excel.Workbook wb);
	public delegate void AppEvents_WorkbookAddinUninstallEventHandler(LateBindingApi.Excel.Workbook wb);
	public delegate void AppEvents_WindowResizeEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.Window wn);
	public delegate void AppEvents_WindowActivateEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.Window wn);
	public delegate void AppEvents_WindowDeactivateEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.Window wn);
	public delegate void AppEvents_SheetFollowHyperlinkEventHandler(COMObject sh, LateBindingApi.Excel.Hyperlink target);
	public delegate void AppEvents_SheetPivotTableUpdateEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable target);
	public delegate void AppEvents_WorkbookPivotTableCloseConnectionEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.PivotTable target);
	public delegate void AppEvents_WorkbookPivotTableOpenConnectionEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.PivotTable target);
	public delegate void AppEvents_WorkbookSyncEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Office.Enums.MsoSyncEventType syncEventType);
	public delegate void AppEvents_WorkbookBeforeXmlImportEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.XmlMap map, string url, bool isRefresh, ref bool cancel);
	public delegate void AppEvents_WorkbookAfterXmlImportEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.XmlMap map, bool isRefresh, LateBindingApi.Excel.Enums.XlXmlImportResult result);
	public delegate void AppEvents_WorkbookBeforeXmlExportEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.XmlMap map, string url, ref bool cancel);
	public delegate void AppEvents_WorkbookAfterXmlExportEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.XmlMap map, string url, LateBindingApi.Excel.Enums.XlXmlExportResult result);
	public delegate void AppEvents_WorkbookRowsetCompleteEventHandler(LateBindingApi.Excel.Workbook wb, string description, string sheet, bool success);
	public delegate void AppEvents_AfterCalculateEventHandler();
	public delegate void AppEvents_SheetPivotTableAfterValueChangeEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, LateBindingApi.Excel.Range targetRange);
	public delegate void AppEvents_SheetPivotTableBeforeAllocateChangesEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void AppEvents_SheetPivotTableBeforeCommitChangesEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd, ref bool cancel);
	public delegate void AppEvents_SheetPivotTableBeforeDiscardChangesEventHandler(COMObject sh, LateBindingApi.Excel.PivotTable targetPivotTable, Int32 valueChangeStart, Int32 valueChangeEnd);
	public delegate void AppEvents_ProtectedViewWindowOpenEventHandler(LateBindingApi.Excel.ProtectedViewWindow pvw);
	public delegate void AppEvents_ProtectedViewWindowBeforeEditEventHandler(LateBindingApi.Excel.ProtectedViewWindow pvw, ref bool cancel);
	public delegate void AppEvents_ProtectedViewWindowBeforeCloseEventHandler(LateBindingApi.Excel.ProtectedViewWindow pvw, LateBindingApi.Excel.Enums.XlProtectedViewCloseReason reason, ref bool cancel);
	public delegate void AppEvents_ProtectedViewWindowResizeEventHandler(LateBindingApi.Excel.ProtectedViewWindow pvw);
	public delegate void AppEvents_ProtectedViewWindowActivateEventHandler(LateBindingApi.Excel.ProtectedViewWindow pvw);
	public delegate void AppEvents_ProtectedViewWindowDeactivateEventHandler(LateBindingApi.Excel.ProtectedViewWindow pvw);
	public delegate void AppEvents_WorkbookAfterSaveEventHandler(LateBindingApi.Excel.Workbook wb, bool success);
	public delegate void AppEvents_WorkbookNewChartEventHandler(LateBindingApi.Excel.Workbook wb, LateBindingApi.Excel.Chart ch);


	#endregion  
	
	#region InterfaceEvents

	public interface AppEvents_Event : IEventBinding 
	{
		event AppEvents_NewWorkbookEventHandler NewWorkbookEvent;
		event AppEvents_SheetSelectionChangeEventHandler SheetSelectionChangeEvent;
		event AppEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent;
		event AppEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent;
		event AppEvents_SheetActivateEventHandler SheetActivateEvent;
		event AppEvents_SheetDeactivateEventHandler SheetDeactivateEvent;
		event AppEvents_SheetCalculateEventHandler SheetCalculateEvent;
		event AppEvents_SheetChangeEventHandler SheetChangeEvent;
		event AppEvents_WorkbookOpenEventHandler WorkbookOpenEvent;
		event AppEvents_WorkbookActivateEventHandler WorkbookActivateEvent;
		event AppEvents_WorkbookDeactivateEventHandler WorkbookDeactivateEvent;
		event AppEvents_WorkbookBeforeCloseEventHandler WorkbookBeforeCloseEvent;
		event AppEvents_WorkbookBeforeSaveEventHandler WorkbookBeforeSaveEvent;
		event AppEvents_WorkbookBeforePrintEventHandler WorkbookBeforePrintEvent;
		event AppEvents_WorkbookNewSheetEventHandler WorkbookNewSheetEvent;
		event AppEvents_WorkbookAddinInstallEventHandler WorkbookAddinInstallEvent;
		event AppEvents_WorkbookAddinUninstallEventHandler WorkbookAddinUninstallEvent;
		event AppEvents_WindowResizeEventHandler WindowResizeEvent;
		event AppEvents_WindowActivateEventHandler WindowActivateEvent;
		event AppEvents_WindowDeactivateEventHandler WindowDeactivateEvent;
		event AppEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent;
		event AppEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent;
		event AppEvents_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnectionEvent;
		event AppEvents_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnectionEvent;
		event AppEvents_WorkbookSyncEventHandler WorkbookSyncEvent;
		event AppEvents_WorkbookBeforeXmlImportEventHandler WorkbookBeforeXmlImportEvent;
		event AppEvents_WorkbookAfterXmlImportEventHandler WorkbookAfterXmlImportEvent;
		event AppEvents_WorkbookBeforeXmlExportEventHandler WorkbookBeforeXmlExportEvent;
		event AppEvents_WorkbookAfterXmlExportEventHandler WorkbookAfterXmlExportEvent;
		event AppEvents_WorkbookRowsetCompleteEventHandler WorkbookRowsetCompleteEvent;
		event AppEvents_AfterCalculateEventHandler AfterCalculateEvent;
		event AppEvents_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent;
		event AppEvents_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent;
		event AppEvents_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent;
		event AppEvents_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent;
		event AppEvents_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent;
		event AppEvents_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent;
		event AppEvents_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent;
		event AppEvents_ProtectedViewWindowResizeEventHandler ProtectedViewWindowResizeEvent;
		event AppEvents_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent;
		event AppEvents_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent;
		event AppEvents_WorkbookAfterSaveEventHandler WorkbookAfterSaveEvent;
		event AppEvents_WorkbookNewChartEventHandler WorkbookNewChartEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	[ComImport, Guid("00024413-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface AppEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1565)]
		void NewWorkbook([In, MarshalAs(UnmanagedType.Interface)] object wb);

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

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1567)]
		void WorkbookOpen([In, MarshalAs(UnmanagedType.Interface)] object wb);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1568)]
		void WorkbookActivate([In, MarshalAs(UnmanagedType.Interface)] object wb);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1569)]
		void WorkbookDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wb);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1570)]
		void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1571)]
		void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] bool saveAsUI, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1572)]
		void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1573)]
		void WorkbookNewSheet([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1574)]
		void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.Interface)] object wb);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1575)]
		void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.Interface)] object wb);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1554)]
		void WindowResize([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1556)]
		void WindowActivate([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1557)]
		void WindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1854)]
		void SheetFollowHyperlink([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2157)]
		void SheetPivotTableUpdate([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2160)]
		void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2161)]
		void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object target);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2289)]
		void WorkbookSync([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] LateBindingApi.Office.Enums.MsoSyncEventType syncEventType);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2290)]
		void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] bool isRefresh, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2291)]
		void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] bool isRefresh, [In] LateBindingApi.Excel.Enums.XlXmlImportResult result);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2292)]
		void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2293)]
		void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] LateBindingApi.Excel.Enums.XlXmlExportResult result);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2611)]
		void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] string description, [In] string sheet, [In] bool success);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2612)]
		void AfterCalculate();

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2895)]
		void SheetPivotTableAfterValueChange([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In, MarshalAs(UnmanagedType.Interface)] object targetRange);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2896)]
		void SheetPivotTableBeforeAllocateChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2897)]
		void SheetPivotTableBeforeCommitChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2898)]
		void SheetPivotTableBeforeDiscardChanges([In, MarshalAs(UnmanagedType.IDispatch)] object sh, [In, MarshalAs(UnmanagedType.Interface)] object targetPivotTable, [In] Int32 valueChangeStart, [In] Int32 valueChangeEnd);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2903)]
		void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.Interface)] object pvw);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2905)]
		void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.Interface)] object pvw, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2906)]
		void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.Interface)] object pvw, [In] LateBindingApi.Excel.Enums.XlProtectedViewCloseReason reason, [In] ref bool cancel);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2908)]
		void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.Interface)] object pvw);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2909)]
		void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.Interface)] object pvw);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2910)]
		void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object pvw);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2911)]
		void WorkbookAfterSave([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] bool success);

		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2912)]
		void WorkbookNewChart([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object ch);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class AppEvents_SinkHelper : SinkHelper, AppEvents
	{
		#region Fields

		private readonly string _riid = "00024413-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public AppEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region AppEvents Members
		
		public void NewWorkbook([In, MarshalAs(UnmanagedType.Interface)] object wb)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			bool isRecieved = _eventBinding.CallEvent("NewWorkbookEvent", paramArray );
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

		public void WorkbookOpen([In, MarshalAs(UnmanagedType.Interface)] object wb)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			bool isRecieved = _eventBinding.CallEvent("WorkbookOpenEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookActivate([In, MarshalAs(UnmanagedType.Interface)] object wb)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			bool isRecieved = _eventBinding.CallEvent("WorkbookActivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wb)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			bool isRecieved = _eventBinding.CallEvent("WorkbookDeactivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookBeforeClose([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray.SetValue(cancel,1);
			bool isRecieved = _eventBinding.CallEvent("WorkbookBeforeCloseEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookBeforeSave([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] bool saveAsUI, [In] ref bool cancel)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = saveAsUI;
			paramArray.SetValue(cancel,2);
			bool isRecieved = _eventBinding.CallEvent("WorkbookBeforeSaveEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookBeforePrint([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray.SetValue(cancel,1);
			bool isRecieved = _eventBinding.CallEvent("WorkbookBeforePrintEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookNewSheet([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.IDispatch)] object sh)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(sh);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new COMObject(_eventClass,sh);
			bool isRecieved = _eventBinding.CallEvent("WorkbookNewSheetEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookAddinInstall([In, MarshalAs(UnmanagedType.Interface)] object wb)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			bool isRecieved = _eventBinding.CallEvent("WorkbookAddinInstallEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookAddinUninstall([In, MarshalAs(UnmanagedType.Interface)] object wb)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			bool isRecieved = _eventBinding.CallEvent("WorkbookAddinUninstallEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WindowResize([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wn);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.Window(_eventClass,wn);
			bool isRecieved = _eventBinding.CallEvent("WindowResizeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WindowActivate([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wn);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.Window(_eventClass,wn);
			bool isRecieved = _eventBinding.CallEvent("WindowActivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object wn)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wn);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.Window(_eventClass,wn);
			bool isRecieved = _eventBinding.CallEvent("WindowDeactivateEvent", paramArray );
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

		public void WorkbookPivotTableCloseConnection([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object target)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("WorkbookPivotTableCloseConnectionEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookPivotTableOpenConnection([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object target)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(target);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.PivotTable(_eventClass,target);
			bool isRecieved = _eventBinding.CallEvent("WorkbookPivotTableOpenConnectionEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookSync([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] LateBindingApi.Office.Enums.MsoSyncEventType syncEventType)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = syncEventType;
			bool isRecieved = _eventBinding.CallEvent("WorkbookSyncEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookBeforeXmlImport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] bool isRefresh, [In] ref bool cancel)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[5];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[2] = url;
			paramArray[3] = isRefresh;
			paramArray.SetValue(cancel,4);
			bool isRecieved = _eventBinding.CallEvent("WorkbookBeforeXmlImportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookAfterXmlImport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] bool isRefresh, [In] LateBindingApi.Excel.Enums.XlXmlImportResult result)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[2] = isRefresh;
			paramArray[3] = result;
			bool isRecieved = _eventBinding.CallEvent("WorkbookAfterXmlImportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookBeforeXmlExport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[2] = url;
			paramArray.SetValue(cancel,3);
			bool isRecieved = _eventBinding.CallEvent("WorkbookBeforeXmlExportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookAfterXmlExport([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object map, [In] string url, [In] LateBindingApi.Excel.Enums.XlXmlExportResult result)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(map);
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.XmlMap(_eventClass,map);
			paramArray[2] = url;
			paramArray[3] = result;
			bool isRecieved = _eventBinding.CallEvent("WorkbookAfterXmlExportEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookRowsetComplete([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] string description, [In] string sheet, [In] bool success)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[4];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = description;
			paramArray[2] = sheet;
			paramArray[3] = success;
			bool isRecieved = _eventBinding.CallEvent("WorkbookRowsetCompleteEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void AfterCalculate()
		{
			bool isRecieved = _eventBinding.CallEvent("AfterCalculateEvent", null );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(null);
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

		public void ProtectedViewWindowOpen([In, MarshalAs(UnmanagedType.Interface)] object pvw)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(pvw);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.ProtectedViewWindow(_eventClass,pvw);
			bool isRecieved = _eventBinding.CallEvent("ProtectedViewWindowOpenEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void ProtectedViewWindowBeforeEdit([In, MarshalAs(UnmanagedType.Interface)] object pvw, [In] ref bool cancel)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(pvw);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.ProtectedViewWindow(_eventClass,pvw);
			paramArray.SetValue(cancel,1);
			bool isRecieved = _eventBinding.CallEvent("ProtectedViewWindowBeforeEditEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void ProtectedViewWindowBeforeClose([In, MarshalAs(UnmanagedType.Interface)] object pvw, [In] LateBindingApi.Excel.Enums.XlProtectedViewCloseReason reason, [In] ref bool cancel)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(pvw);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new LateBindingApi.Excel.ProtectedViewWindow(_eventClass,pvw);
			paramArray[1] = reason;
			paramArray.SetValue(cancel,2);
			bool isRecieved = _eventBinding.CallEvent("ProtectedViewWindowBeforeCloseEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void ProtectedViewWindowResize([In, MarshalAs(UnmanagedType.Interface)] object pvw)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(pvw);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.ProtectedViewWindow(_eventClass,pvw);
			bool isRecieved = _eventBinding.CallEvent("ProtectedViewWindowResizeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void ProtectedViewWindowActivate([In, MarshalAs(UnmanagedType.Interface)] object pvw)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(pvw);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.ProtectedViewWindow(_eventClass,pvw);
			bool isRecieved = _eventBinding.CallEvent("ProtectedViewWindowActivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void ProtectedViewWindowDeactivate([In, MarshalAs(UnmanagedType.Interface)] object pvw)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(pvw);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Excel.ProtectedViewWindow(_eventClass,pvw);
			bool isRecieved = _eventBinding.CallEvent("ProtectedViewWindowDeactivateEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookAfterSave([In, MarshalAs(UnmanagedType.Interface)] object wb, [In] bool success)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = success;
			bool isRecieved = _eventBinding.CallEvent("WorkbookAfterSaveEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		public void WorkbookNewChart([In, MarshalAs(UnmanagedType.Interface)] object wb, [In, MarshalAs(UnmanagedType.Interface)] object ch)
		{
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(ch);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Excel.Workbook(_eventClass,wb);
			paramArray[1] = new LateBindingApi.Excel.Chart(_eventClass,ch);
			bool isRecieved = _eventBinding.CallEvent("WorkbookNewChartEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		#endregion
	}
    
	#endregion
}

