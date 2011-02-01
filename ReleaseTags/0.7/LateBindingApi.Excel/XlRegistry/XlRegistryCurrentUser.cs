using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.XlRegistry
{

    public static class XlRegistryCurrentUser
    {       
        #region Constants

        private const string _rootKey = @"SOFTWARE\Microsoft\Office\Excel";
        
        #endregion
        
        #region Fields

        private static XlRegistryKey _key;

        private static XlRegistryEntries _entries;
        
        #endregion

        #region Properties

        public static bool Exists
        {
            get
            {
                bool retValue = false;
                RegistryKey rk = Registry.CurrentUser.OpenSubKey(_rootKey, false);
                if (rk != null)
                {
                    rk.Close();
                    retValue = true;
                }

                return retValue;
            }
        }

        public static XlRegistryKey Key
        {
            get
            {
                if (null == _key)
                {
                    _key = new XlRegistryKey(LateBindingApi.Excel.Enums.XlRegistryType.HKEY_CURRENT_USER, _rootKey);
                }
                return _key;
            }
        }

        public static XlRegistryEntries Entries
        {
            get 
            {
                if (null == _entries)
                {
                    _entries = new XlRegistryEntries(XlRegistryType.HKEY_CURRENT_USER, _rootKey);
                }
                return _entries;
            }
        }

        #endregion
    }

}
