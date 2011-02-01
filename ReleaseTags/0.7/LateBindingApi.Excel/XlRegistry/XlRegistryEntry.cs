using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;


namespace LateBindingApi.Excel.XlRegistry
{
    public class XlRegistryEntry
    {
        
        #region Fields

        private string             _valueName;
        private object             _value;
        private RegistryValueKind  _valueType;

        #endregion

        #region Properties

        public object Value
        {
            get
            {
                return _value;
            }
        }

        public string Name
        {
            get 
            {
                return _valueName; 
            }
        }

        public RegistryValueKind ValueType
        {
            get
            {
                return _valueType;
            }
        }

        #endregion

        #region Construction

        internal XlRegistryEntry(string ValueName, object Value, RegistryValueKind ValueType)
        {

            _valueName  = ValueName;
            _value      = Value;
            _valueType  = ValueType;
 
            /*
               switch (rvk)
               { 
                   case RegistryValueKind.Binary:

                       break;
                   case RegistryValueKind.DWord:

                       break;
                   case RegistryValueKind.ExpandString:

                       break;
                   case RegistryValueKind.MultiString:

                       break;
                   case RegistryValueKind.QWord:

                       break;
                   case RegistryValueKind.String:
                      
                       break;
                   case RegistryValueKind.Unknown :

                       break;
               }
               */
        }

        #endregion
    }
}
