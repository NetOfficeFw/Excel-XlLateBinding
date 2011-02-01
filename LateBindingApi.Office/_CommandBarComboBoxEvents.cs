using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.Office
{
	#region Event Delegates

	public delegate void _CommandBarComboBoxEvents_ChangeEventHandler(LateBindingApi.Office.CommandBarComboBox ctrl);


	#endregion  
	
	#region InterfaceEvents

	public interface _CommandBarComboBoxEvents_Event : IEventBinding 
	{
		event _CommandBarComboBoxEvents_ChangeEventHandler ChangeEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	[ComImport, Guid("000C0354-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface _CommandBarComboBoxEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
		void Change([In, MarshalAs(UnmanagedType.Interface)] object ctrl);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class _CommandBarComboBoxEvents_SinkHelper : SinkHelper, _CommandBarComboBoxEvents
	{
		#region Fields

		private readonly string _riid = "000C0354-0000-0000-C000-000000000046";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public _CommandBarComboBoxEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region _CommandBarComboBoxEvents Members
		
		public void Change([In, MarshalAs(UnmanagedType.Interface)] object ctrl)
        {
            if (true == _eventClass.IsDisposed)
            {
                Marshal.ReleaseComObject(ctrl);
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = new LateBindingApi.Office.CommandBarComboBox(_eventClass,ctrl);
			bool isRecieved = _eventBinding.CallEvent("ChangeEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		#endregion
	}
    
	#endregion
}

