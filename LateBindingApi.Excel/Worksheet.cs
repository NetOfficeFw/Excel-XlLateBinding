//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class Worksheet : _Worksheet, DocEvents_Event
	{
		#region Fields

		DocEvents_SinkHelper _sinkHelper;

		#endregion

		#region Construction

		public Worksheet(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_sinkHelper = new DocEvents_SinkHelper(this);
		}

		public Worksheet(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_sinkHelper = new DocEvents_SinkHelper(this);
		}

		public Worksheet(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Worksheet()
		{
			CreateFromProgId("Excel.Worksheet");
			_sinkHelper = new DocEvents_SinkHelper(this);
		}

		public Worksheet(string progId)
		{
			CreateFromProgId(progId);
			_sinkHelper = new DocEvents_SinkHelper(this);
		}

		#endregion

        #region DocEvents_Event Members
        
		#pragma warning disable
		public event DocEvents_SelectionChangeEventHandler SelectionChangeEvent;
		public event DocEvents_BeforeDoubleClickEventHandler BeforeDoubleClickEvent;
		public event DocEvents_BeforeRightClickEventHandler BeforeRightClickEvent;
		public event DocEvents_ActivateEventHandler ActivateEvent;
		public event DocEvents_DeactivateEventHandler DeactivateEvent;
		public event DocEvents_CalculateEventHandler CalculateEvent;
		public event DocEvents_ChangeEventHandler ChangeEvent;
		public event DocEvents_FollowHyperlinkEventHandler FollowHyperlinkEvent;
		public event DocEvents_PivotTableUpdateEventHandler PivotTableUpdateEvent;
		public event DocEvents_PivotTableAfterValueChangeEventHandler PivotTableAfterValueChangeEvent;
		public event DocEvents_PivotTableBeforeAllocateChangesEventHandler PivotTableBeforeAllocateChangesEvent;
		public event DocEvents_PivotTableBeforeCommitChangesEventHandler PivotTableBeforeCommitChangesEvent;
		public event DocEvents_PivotTableBeforeDiscardChangesEventHandler PivotTableBeforeDiscardChangesEvent;
		public event DocEvents_PivotTableChangeSyncEventHandler PivotTableChangeSyncEvent;

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
				_sinkHelper.Dispose();
			}
        }
        
        #endregion 
	}
}
