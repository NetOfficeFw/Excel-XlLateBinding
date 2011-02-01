//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class Application : _Application, AppEvents_Event
	{
		#region Fields

		AppEvents_SinkHelper _sinkHelper;

		#endregion

		#region Construction

		public Application(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_sinkHelper = new AppEvents_SinkHelper(this);
		}

		public Application(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_sinkHelper = new AppEvents_SinkHelper(this);
		}

		public Application(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Application()
		{
			CreateFromProgId("Excel.Application");
			_sinkHelper = new AppEvents_SinkHelper(this);
		}

		public Application(string progId)
		{
			CreateFromProgId(progId);
			_sinkHelper = new AppEvents_SinkHelper(this);
		}

		#endregion

        #region AppEvents_Event Members
        
		#pragma warning disable
		public event AppEvents_NewWorkbookEventHandler NewWorkbookEvent;
		public event AppEvents_SheetSelectionChangeEventHandler SheetSelectionChangeEvent;
		public event AppEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClickEvent;
		public event AppEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClickEvent;
		public event AppEvents_SheetActivateEventHandler SheetActivateEvent;
		public event AppEvents_SheetDeactivateEventHandler SheetDeactivateEvent;
		public event AppEvents_SheetCalculateEventHandler SheetCalculateEvent;
		public event AppEvents_SheetChangeEventHandler SheetChangeEvent;
		public event AppEvents_WorkbookOpenEventHandler WorkbookOpenEvent;
		public event AppEvents_WorkbookActivateEventHandler WorkbookActivateEvent;
		public event AppEvents_WorkbookDeactivateEventHandler WorkbookDeactivateEvent;
		public event AppEvents_WorkbookBeforeCloseEventHandler WorkbookBeforeCloseEvent;
		public event AppEvents_WorkbookBeforeSaveEventHandler WorkbookBeforeSaveEvent;
		public event AppEvents_WorkbookBeforePrintEventHandler WorkbookBeforePrintEvent;
		public event AppEvents_WorkbookNewSheetEventHandler WorkbookNewSheetEvent;
		public event AppEvents_WorkbookAddinInstallEventHandler WorkbookAddinInstallEvent;
		public event AppEvents_WorkbookAddinUninstallEventHandler WorkbookAddinUninstallEvent;
		public event AppEvents_WindowResizeEventHandler WindowResizeEvent;
		public event AppEvents_WindowActivateEventHandler WindowActivateEvent;
		public event AppEvents_WindowDeactivateEventHandler WindowDeactivateEvent;
		public event AppEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent;
		public event AppEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdateEvent;
		public event AppEvents_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnectionEvent;
		public event AppEvents_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnectionEvent;
		public event AppEvents_WorkbookSyncEventHandler WorkbookSyncEvent;
		public event AppEvents_WorkbookBeforeXmlImportEventHandler WorkbookBeforeXmlImportEvent;
		public event AppEvents_WorkbookAfterXmlImportEventHandler WorkbookAfterXmlImportEvent;
		public event AppEvents_WorkbookBeforeXmlExportEventHandler WorkbookBeforeXmlExportEvent;
		public event AppEvents_WorkbookAfterXmlExportEventHandler WorkbookAfterXmlExportEvent;
		public event AppEvents_WorkbookRowsetCompleteEventHandler WorkbookRowsetCompleteEvent;
		public event AppEvents_AfterCalculateEventHandler AfterCalculateEvent;
		public event AppEvents_SheetPivotTableAfterValueChangeEventHandler SheetPivotTableAfterValueChangeEvent;
		public event AppEvents_SheetPivotTableBeforeAllocateChangesEventHandler SheetPivotTableBeforeAllocateChangesEvent;
		public event AppEvents_SheetPivotTableBeforeCommitChangesEventHandler SheetPivotTableBeforeCommitChangesEvent;
		public event AppEvents_SheetPivotTableBeforeDiscardChangesEventHandler SheetPivotTableBeforeDiscardChangesEvent;
		public event AppEvents_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent;
		public event AppEvents_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent;
		public event AppEvents_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent;
		public event AppEvents_ProtectedViewWindowResizeEventHandler ProtectedViewWindowResizeEvent;
		public event AppEvents_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent;
		public event AppEvents_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent;
		public event AppEvents_WorkbookAfterSaveEventHandler WorkbookAfterSaveEvent;
		public event AppEvents_WorkbookNewChartEventHandler WorkbookNewChartEvent;

		#pragma warning restore
		
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool CallEvent(string name, object[] paramArray)
        {
            Type thisType = this.GetType();

            MulticastDelegate eventDelegate = (MulticastDelegate)thisType.GetField(
												name, 
												System.Reflection.BindingFlags.Instance|
												System.Reflection.BindingFlags.NonPublic).GetValue(this);
            
            if(null!=eventDelegate)
            {            
				Delegate[] delegates = eventDelegate.GetInvocationList();
				
				foreach (Delegate invocation in delegates)
					invocation.Method.Invoke(invocation.Target, paramArray);
	                    
				return (delegates.Length > 0);
            }
            else
				return false;
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeSinkHelper()
        {
            if (null != _sinkHelper)
			{
                _sinkHelper.Dispose();
                _sinkHelper = null;
			}
        }
        
        #endregion 
	}
}