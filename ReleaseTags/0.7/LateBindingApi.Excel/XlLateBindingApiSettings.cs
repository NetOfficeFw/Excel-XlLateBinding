using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;

namespace LateBindingApi.Excel
{
    /// <summary>
    /// Behaviour Settings for the Framework
    /// </summary>
    public static class XlLateBindingApiSettings
    {
        #region Imports

        [DllImport("ole32.dll", ExactSpelling = true)]
        private static extern int CoRegisterMessageFilter(IntPtr newFilter, ref IntPtr oldMsgFilter);

        #endregion
        
        #region Constants

        // Default Thread Culture to Excel Application
        private const string _DefaultCulture = "en-US";
        
        #endregion

        #region Fields

        private static CultureInfo  _cultureInfo;
        private static bool         _eventsEnabled;
        private static bool         _messageFilterEnabled;
        private static IntPtr       _messageFilter;
 
        #endregion

        #region Properties

        /// <summary>
        /// Assembly Version
        /// </summary>
        public static Version FrameworkVersion
        {

            get
            {
                return Assembly.GetExecutingAssembly().GetName().Version;
            }

        }

        /// <summary>
        /// The Default ThreadLCID for XlThreadCulture at initialize on demand
        /// </summary>
        public static string DefaultThreadLCID
        {

            get
            {
                return _DefaultCulture;
            }

        }

        /// <summary>
        /// Used Thread Culture given in the Invoke Calls, default is DefaultThreadLCID
        /// </summary>
        public static CultureInfo XlThreadCulture
        {

            get
            {
                try
                {
                    if (null == _cultureInfo)
                        _cultureInfo = CultureInfo.GetCultureInfo(XlLateBindingApiSettings.DefaultThreadLCID);
                }
                catch (Exception ex)
                {
                    throw (new Exception(string.Format("GetCultureInfo {0} failed.", XlLateBindingApiSettings.DefaultThreadLCID), ex));
                }
                finally
                {
                    if (null == _cultureInfo)
                        throw (new Exception(string.Format("GetCultureInfo {0} failed.", XlLateBindingApiSettings.DefaultThreadLCID)));
                }

                return _cultureInfo;
            }
            set
            {
                _cultureInfo = value;
            }

        }

        /// <summary>
        /// Get or set the Event support, default is false 
        /// </summary>
        public static bool EventsEnabled
        {
            get
            {
                return _eventsEnabled;
            }
            set
            {
                _eventsEnabled = value;
            }
        }
      
        /// <summary>
        /// Get or set the Message Filter enabled.
        /// The MessageFilter suspress any exceptional dialog messages, specialy the "Excel ist waiting for another OLE Task" dialog
        /// </summary>
        public static bool MessageFilterEnabled
        {
            get
            {
                return _messageFilterEnabled;
            }
            set
            {
                if( (value == true) && (IntPtr.Zero ==_messageFilter))
                {
                    CoRegisterMessageFilter((IntPtr)0, ref _messageFilter);
                }
                else if ((value == false) && (IntPtr.Zero != _messageFilter))
                {
                    IntPtr filter = IntPtr.Zero;
                    CoRegisterMessageFilter(_messageFilter, ref filter);
                }
                _messageFilterEnabled = value;
            }
        }

        #endregion
    }
}

