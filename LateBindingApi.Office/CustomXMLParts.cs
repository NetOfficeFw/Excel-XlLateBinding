//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class CustomXMLParts : _CustomXMLParts, _CustomXMLPartsEvents_Event
	{
		#region Fields

		_CustomXMLPartsEvents_SinkHelper _sinkHelper;

		#endregion

		#region Construction

		public CustomXMLParts(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_sinkHelper = new _CustomXMLPartsEvents_SinkHelper(this);
		}

		public CustomXMLParts(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_sinkHelper = new _CustomXMLPartsEvents_SinkHelper(this);
		}

		public CustomXMLParts(COMObject replacedObject) : base(replacedObject)
		{
		}

		public CustomXMLParts()
		{
			CreateFromProgId("Office.CustomXMLParts");
			_sinkHelper = new _CustomXMLPartsEvents_SinkHelper(this);
		}

		public CustomXMLParts(string progId)
		{
			CreateFromProgId(progId);
			_sinkHelper = new _CustomXMLPartsEvents_SinkHelper(this);
		}

		#endregion

        #region _CustomXMLPartsEvents_Event Members
        
		#pragma warning disable
		public event _CustomXMLPartsEvents_PartAfterAddEventHandler PartAfterAddEvent;
		public event _CustomXMLPartsEvents_PartBeforeDeleteEventHandler PartBeforeDeleteEvent;
		public event _CustomXMLPartsEvents_PartAfterLoadEventHandler PartAfterLoadEvent;

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
