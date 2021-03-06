//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class MsoEnvelope : IMsoEnvelopeVB, IMsoEnvelopeVBEvents_Event
	{
		#region Fields

		IMsoEnvelopeVBEvents_SinkHelper _sinkHelper;

		#endregion

		#region Construction

		public MsoEnvelope(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			_sinkHelper = new IMsoEnvelopeVBEvents_SinkHelper(this);
		}

		public MsoEnvelope(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			_sinkHelper = new IMsoEnvelopeVBEvents_SinkHelper(this);
		}

		public MsoEnvelope(COMObject replacedObject) : base(replacedObject)
		{
		}

		public MsoEnvelope()
		{
			CreateFromProgId("Office.MsoEnvelope");
			_sinkHelper = new IMsoEnvelopeVBEvents_SinkHelper(this);
		}

		public MsoEnvelope(string progId)
		{
			CreateFromProgId(progId);
			_sinkHelper = new IMsoEnvelopeVBEvents_SinkHelper(this);
		}

		#endregion

        #region IMsoEnvelopeVBEvents_Event Members
        
		#pragma warning disable
		public event IMsoEnvelopeVBEvents_EnvelopeShowEventHandler EnvelopeShowEvent;
		public event IMsoEnvelopeVBEvents_EnvelopeHideEventHandler EnvelopeHideEvent;

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
