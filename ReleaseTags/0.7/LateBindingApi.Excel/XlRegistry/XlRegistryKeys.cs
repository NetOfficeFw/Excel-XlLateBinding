using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.XlRegistry
{
    
    public class XlRegistryKeys
    {

        #region Fields

        private string                _key;
        private List<XlRegistryKey>   _list;
        private XlRegistryType        _regType;

        #endregion

        #region Properties

        public XlRegistryType RegistryType 
        {

            get 
            {
                return _regType;
            }

        }

        public int Count
        {
            get
            {
                return _list.Count;
            }
        }

        public XlRegistryKey this[int index]
        {
            get
            {
                return _list[index];
            }
        }

        public XlRegistryKey this[string Name]
        {
            get
            {
                int iCount = Count;
                for (int i = 1; i <= iCount; i++)
                {
                    XlRegistryKey entry = this[i - 1];
                    if (Name.Equals(entry.Name, StringComparison.CurrentCultureIgnoreCase) == true)
                    {
                        return entry;
                    }
                }

                throw (new IndexOutOfRangeException("RegistryEntry " + Name + " not found."));
            }
        }

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            int iCount = this.Count;
            XlRegistryKey[] res_keys = new XlRegistryKey[iCount];

            for (int i = 0; i < iCount; i++)
                res_keys[i] = this[i];

            for (int i = 0; i < res_keys.Length; i++)
            {
                yield return res_keys[i];
            }

        }

        #endregion

        #region Construction

        internal XlRegistryKeys(XlRegistryType regType, string Key)
        {
           
            _key    = Key;
            _list   = new List<XlRegistryKey>();
            _regType = regType;

            RegistryKey rk;

            if (_regType == XlRegistryType.HKEY_CURRENT_USER)
            {
                rk = Registry.CurrentUser.OpenSubKey(_key, false);
            }
            else
            {
                rk = Registry.LocalMachine.OpenSubKey(_key, false);
            }

            if (null != rk)
            { 
                string[] Subkeys = rk.GetSubKeyNames();
                foreach (string sKey in Subkeys)
                {
                    XlRegistryKey NewKey = new XlRegistryKey(regType, _key + "\\" + sKey);
                    _list.Add(NewKey);

                }
                rk.Close();
            }
        }

        #endregion

    }

}
