//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class References : _References, _dispReferences_Events_Event
	{
		#region Fields

		_dispReferences_Events_SinkHelper _sinkHelper;

		#endregion

		#region Construction

		public References(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_sinkHelper = new _dispReferences_Events_SinkHelper(this);
		}

		public References(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_sinkHelper = new _dispReferences_Events_SinkHelper(this);
		}

		public References(COMObject replacedObject) : base(replacedObject)
		{
		}

		public References()
		{
			CreateFromProgId("VBIDE.References");
			_sinkHelper = new _dispReferences_Events_SinkHelper(this);
		}

		public References(string progId)
		{
			CreateFromProgId(progId);
			_sinkHelper = new _dispReferences_Events_SinkHelper(this);
		}

		#endregion

        #region _dispReferences_Events_Event Members
        
		#pragma warning disable
		public event _dispReferences_Events_ItemAddedEventHandler ItemAddedEvent;
		public event _dispReferences_Events_ItemRemovedEventHandler ItemRemovedEvent;

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
