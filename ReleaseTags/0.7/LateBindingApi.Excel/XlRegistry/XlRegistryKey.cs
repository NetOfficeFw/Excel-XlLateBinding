using System;
using System.Collections.Generic;
using System.Text;

using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.XlRegistry
{
    public class XlRegistryKey
    {
        #region Fields

        private string          _Key;
        private string          _Name;
        private XlRegistryType  _regType;

        private XlRegistryEntries _entries;
        private XlRegistryKeys    _subKeys;

        #endregion

        #region Properties

        public XlRegistryType RegistryType
        {

            get
            {
                return _regType;
            }

        }

        public string Name
        {
            get
            {
                return _Name;
            }
        }

       
        public XlRegistryEntries Entries
        {
            get
            {
                return _entries;
            }
        }
       

        public XlRegistryKeys Keys
        {
            get
            {
                return _subKeys;
            }
        }
       
        #endregion

        #region Construction

        internal XlRegistryKey(XlRegistryType regType, string RootKey)
        {
              
            _regType = regType;
            _Key = RootKey;
            _Name = RootKey.Substring(RootKey.LastIndexOf(@"\",StringComparison.Ordinal)+1);
            
            _entries = new XlRegistryEntries(_regType,_Key);
            _subKeys = new XlRegistryKeys(regType, _Key);

        }

        #endregion
    }
}
