using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using LateBindingApi.Core;

namespace LateBindingApi.stdole
{
	#region Event Delegates

	public delegate void FontEvents_FontChangedEventHandler(string propertyName);


	#endregion  
	
	#region InterfaceEvents

	public interface FontEvents_Event : IEventBinding 
	{
		event FontEvents_FontChangedEventHandler FontChangedEvent;

	}
	
	#endregion
	
	#region SinkPoint Interface
	
	[SupportByLibrary("OLE")]
	[ComImport, Guid("4EF6100A-AF88-11D0-9846-00C04FC29993"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface FontEvents
	{
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(9)]
		void FontChanged([In] string propertyName);

    
	}
    	
	#endregion
	
	#region SinkHelper
	
	[ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class FontEvents_SinkHelper : SinkHelper, FontEvents
	{
		#region Fields

		private readonly string _riid = "4EF6100A-AF88-11D0-9846-00C04FC29993";
		private IEventBinding	_eventBinding;
        private COMObject		_eventClass;
        
		#endregion
        
		#region Construction

		public FontEvents_SinkHelper(COMObject eventClass): base(eventClass)
		{
			_eventClass = eventClass;
			_eventBinding = (IEventBinding)eventClass;
			SetupEventBinding(_riid);
		}

		#endregion
        
		#region FontEvents Members
		
		public void FontChanged([In] string propertyName)
		{
            if (true == _eventClass.IsDisposed)
            {
                return;
            }

			object[] paramArray = new object[1];
			paramArray[0] = propertyName;
			bool isRecieved = _eventBinding.CallEvent("FontChangedEvent", paramArray );
			if (false == isRecieved)
				Invoker.ReleaseParamArray(paramArray);
		}

		#endregion
	}
    
	#endregion
}

