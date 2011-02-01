using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Office
{
	#region Event Delegates

	public delegate void _CommandBarButtonEvents_ClickEventHandler(LateBindingApi.Office.CommandBarButton ctrl, ref bool cancelDefault);


	#endregion  
	
	#region InterfaceEvents

	public interface _CommandBarButtonEvents_Event : IEventBinding 
	{
		event _CommandBarButtonEvents_ClickEventHandler ClickEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	[ComImport, Guid("000C0351-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarButtonEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Click([In, MarshalAs(UnmanagedType.Interface)] object ctrl, [In] ref bool cancelDefault);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CommandBarButtonEvents_SinkHelper : SinkHelper, _CommandBarButtonEvents
	{
		#region Fields

		private readonly string _riid = "000C0351-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _CommandBarButtonEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _CommandBarButtonEvents Members
		
		public void Click([In, MarshalAs(UnmanagedType.Interface)] object ctrl, [In] ref bool cancelDefault)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(ctrl);
                return;
            }

			object[] paramArray = new object[2];
			paramArray[0] = new LateBindingApi.Office.CommandBarButton(_eventClass,ctrl);
			paramArray.SetValue(cancelDefault,1);
			bool isRecieved = _eventBinding.CallEvent("ClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		#endregion
	}
    
	#endregion
}

