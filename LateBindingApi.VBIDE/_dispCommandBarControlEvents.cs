using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.VBIDE
{
	#region Event Delegates

	public delegate void _dispCommandBarControlEvents_ClickEventHandler(COMObject commandBarControl, ref bool handled, ref bool cancelDefault);


	#endregion  
	
	#region InterfaceEvents

	public interface _dispCommandBarControlEvents_Event : IEventBinding 
	{
		event _dispCommandBarControlEvents_ClickEventHandler ClickEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("VBE")]
	[ComImport, Guid("0002E131-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _dispCommandBarControlEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] ref bool handled, [In] ref bool cancelDefault);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _dispCommandBarControlEvents_SinkHelper : SinkHelper, _dispCommandBarControlEvents
	{
		#region Fields

		private readonly string _riid = "0002E131-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _dispCommandBarControlEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _dispCommandBarControlEvents Members
		
		public void Click([In, MarshalAs(UnmanagedType.IDispatch)] object commandBarControl, [In] ref bool handled, [In] ref bool cancelDefault)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(commandBarControl);
                return;
            }

			object[] paramArray = new object[3];
			paramArray[0] = new COMObject(_eventClass,commandBarControl);
			paramArray.SetValue(handled,1);
			paramArray.SetValue(cancelDefault,2);
			bool isRecieved = _eventBinding.CallEvent("ClickEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}



		#endregion
	}
    
	#endregion
}

