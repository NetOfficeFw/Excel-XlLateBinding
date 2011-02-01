using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;

using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.XlRegistry
{

    public class XlRegistryEntries
    {

        #region Fields

        private string _key = "";
     
        private List<XlRegistryEntry> _list = null;

        #endregion

        #region Properties

        public int Count
        {
            get 
            {
                return _list.Count;
            }
        }

        public XlRegistryEntry this[int index]
        {
            get 
            {
                return _list[index];
            }
        }

        public XlRegistryEntry this[string name]
        {
            get
            {
                int iCount = Count;
                for (int i = 1; i <= iCount; i++)
                {
                    XlRegistryEntry entry = this[i-1];
                    if (name.Equals(entry.Name, StringComparison.CurrentCultureIgnoreCase) == true)
                    {
                        return entry;
                    }
                }

                throw (new IndexOutOfRangeException("RegistryEntry " + name + " not found."));
            }
        }

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            int iCount = Count;
            XlRegistryEntry[] res_entries = new XlRegistryEntry[iCount];

            for (int i = 0; i < iCount; i++)
                res_entries[i] = this[i];

            for (int i = 0; i < res_entries.Length; i++)
            {
                yield return res_entries[i];
            }

        }
 
        #endregion

        #region Construction

        internal XlRegistryEntries(XlRegistryType regType,  string Key)
        {

            _key    = Key;
            _list   = new List<XlRegistryEntry>();

            RegistryKey rk = null;
            
            if(regType == XlRegistryType.HKEY_CURRENT_USER)
                rk =  Registry.CurrentUser.OpenSubKey(_key, false);
            else
                rk = Registry.LocalMachine.OpenSubKey(_key, false);

            if (null != rk)
            { 
                string[] Values = rk.GetValueNames();
                foreach (string Value in Values)
                {
                    XlRegistryEntry Entry = null;
                    RegistryValueKind rvk = rk.GetValueKind(Value);
                    object o = rk.GetValue(Value);
                    Entry = new XlRegistryEntry(Value, o, rvk);
                    _list.Add(Entry); 
                }
                
                rk.Close();
            }
        }

        #endregion

    }

}
