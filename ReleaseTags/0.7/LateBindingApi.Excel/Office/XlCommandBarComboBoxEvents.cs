using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Security.Permissions;

namespace LateBindingApi.Excel.Office
{  
    #region Event Delegates

    public delegate void CommandBarComboBoxEvents_ChangeEventHandler([In, MarshalAs(UnmanagedType.Interface)] XlCommandBarComboBox Ctrl);
        
    #endregion
 
    [ComImport, Guid("000C0354-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
    public interface ICommandBarComboBoxEvents
    {
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void Change([In, MarshalAs(UnmanagedType.Interface)] object Ctrl);
    }

    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
    public class XlCommandBarComboBoxEvents : ICommandBarComboBoxEvents
    {
        #region Fields

        private XlCommandBarComboBox  _comboBox;
        private IConnectionPoint _connectionPoint;
        private int _connectionCookie;

        #endregion
        
        #region Construction

        public XlCommandBarComboBoxEvents()
        {
        }
        
        #endregion

        #region ICommandBarComboBoxEvents Members

        public void Change(object Ctrl)
        {
            _comboBox.RaiseChangeEvent(Ctrl);
        }

        #endregion

        #region SetupEventBinding
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        public void SetupEventBinding(XlCommandBarComboBox comboBox)
        {
            if (true == XlLateBindingApiSettings.EventsEnabled)
            {
                _comboBox = comboBox;
                IConnectionPointContainer connectionPointContainer = (IConnectionPointContainer)comboBox.COMReference;
                Guid guid = new Guid("{000C0354-0000-0000-C000-000000000046}");
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
