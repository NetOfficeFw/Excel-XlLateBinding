using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;


namespace LateBindingApi.Excel.Office
{  
    #region Event Delegates

    public delegate void CommandBarsEvents_OnUpdateEventHandler();
      
    #endregion

    [ComImport, Guid("000C0352-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface ICommandBarsEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void OnUpdate();
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class XlCommandBarsEvents : ICommandBarsEvents
    {
        #region Fields

        private XlCommandBars  _commandBar;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion
        
        #region Construction

        public XlCommandBarsEvents()
        {
        }
        
        #endregion

        #region ICommandBarsEvents Members

        public void OnUpdate()
        {
            _commandBar.RaiseOnUpdateEvent();
        }

        #endregion

        #region SetupEventBinding

        public void SetupEventBinding(XlCommandBars commandBar)
        {
            if (true == XlLateBindingApiSettings.EventsEnabled)
            {
                _commandBar = commandBar;
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)commandBar.COMReference;
                Guid guid = new Guid("{000C0352-0000-0000-C000-000000000046}");
                connectionPointContainer.FindConnectionPoint(ref guid, out _connectionPoint);
                _connectionPoint.Advise(this, out _connectionCookie);
            }
        }

        public void RemoveEventBinding()
        {
            if (_connectionCookie != 0)
            {
                _connectionPoint.Unadvise(_connectionCookie);
                Marshal.ReleaseComObject(_connectionPoint);
                _connectionPoint = null;
                _connectionCookie = 0;
            }
        }

        #endregion
    }
}
