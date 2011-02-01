//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class CustomTaskPane : _CustomTaskPane, _CustomTaskPaneEvents_Event
	{
		#region Fields

		_CustomTaskPaneEvents_SinkHelper _sinkHelper;

		#endregion

		#region Construction

		public CustomTaskPane(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_sinkHelper = new _CustomTaskPaneEvents_SinkHelper(this);
		}

		public CustomTaskPane(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_sinkHelper = new _CustomTaskPaneEvents_SinkHelper(this);
		}

		public CustomTaskPane(COMObject replacedObject) : base(replacedObject)
		{
		}

		public CustomTaskPane()
		{
			CreateFromProgId("Office.CustomTaskPane");
			_sinkHelper = new _CustomTaskPaneEvents_SinkHelper(this);
		}

		public CustomTaskPane(string progId)
		{
			CreateFromProgId(progId);
			_sinkHelper = new _CustomTaskPaneEvents_SinkHelper(this);
		}

		#endregion

        #region _CustomTaskPaneEvents_Event Members
        
		#pragma warning disable
		public event _CustomTaskPaneEvents_VisibleStateChangeEventHandler VisibleStateChangeEvent;
		public event _CustomTaskPaneEvents_DockPositionStateChangeEventHandler DockPositionStateChangeEvent;

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