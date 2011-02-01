using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Security.Permissions;

namespace LateBindingApi.Excel.Office
{
    #region Event Delegates
    
    public delegate void CommandBarButtonEvents_ClickEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlCommandBarButton Ctrl, [In, Out] ref bool CancelDefault);
 
    #endregion
 
    [ComImport, Guid("000C0351-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface ICommandBarButtonEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void Click([In, MarshalAs(UnmanagedType.Interface)] object Ctrl, [In, Out] ref bool CancelDefault);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class XlCommandBarButtonEvents : ICommandBarButtonEvents
    {
        
        #region Fields

        private XlCommandBarButton  _button;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion
        
        #region Construction

        public XlCommandBarButtonEvents()
        {
        }
        
        #endregion

        #region ICommandBarButtonEvents Members

        public void Click(object Ctrl, ref bool CancelDefault)
        {
            _button.RaiseClickEvent(Ctrl, ref CancelDefault);
        }

        #endregion

        #region SetupEventBinding

        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        public void SetupEventBinding(XlCommandBarButton  button)
        {

            if (true == XlLateBindingApiSettings.EventsEnabled)
            {
                _button = button;
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)button.COMReference;
                Guid guid = new Guid("{000C0351-0000-0000-C000-000000000046}");
                connectionPointContainer.FindConnectionPoint(ref guid, out _connectionPoint);
                _connectionPoint.Advise(this, out _connectionCookie);
            }

        }

        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
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
