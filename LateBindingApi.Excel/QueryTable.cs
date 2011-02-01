//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class QueryTable : _QueryTable, RefreshEvents_Event
	{
		#region Fields

		RefreshEvents_SinkHelper _sinkHelper;

		#endregion

		#region Construction

		public QueryTable(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_sinkHelper = new RefreshEvents_SinkHelper(this);
		}

		public QueryTable(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_sinkHelper = new RefreshEvents_SinkHelper(this);
		}

		public QueryTable(COMObject replacedObject) : base(replacedObject)
		{
		}

		public QueryTable()
		{
			CreateFromProgId("Excel.QueryTable");
			_sinkHelper = new RefreshEvents_SinkHelper(this);
		}

		public QueryTable(string progId)
		{
			CreateFromProgId(progId);
			_sinkHelper = new RefreshEvents_SinkHelper(this);
		}

		#endregion

        #region RefreshEvents_Event Members
        
		#pragma warning disable
		public event RefreshEvents_BeforeRefreshEventHandler BeforeRefreshEvent;
		public event RefreshEvents_AfterRefreshEventHandler AfterRefreshEvent;

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