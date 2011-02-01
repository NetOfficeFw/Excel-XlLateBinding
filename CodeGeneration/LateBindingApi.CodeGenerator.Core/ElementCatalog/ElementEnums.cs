using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.CodeGenerator.Core
{
    public class ElementEnums
    {
        #region Fields

        Dictionary<string, ElementEnum> _enums = new Dictionary<string, ElementEnum>();
        
        #endregion

        #region Properties

        public int Count
        {
            get 
            {
                return _enums.Count;
            }
        }

        #endregion

        #region Methods

        internal void Add(COMComponent[] components)        
        {
            foreach (COMComponent component in components)
            {
                foreach (COMEnum enumItem in component.Enums)
                {
                    ElementEnum newEnumElement = new ElementEnum(enumItem);
                    _enums.Add(enumItem.Name, newEnumElement); 
                }
            }        
        }

        internal void Clear()
        {
            _enums.Clear();
        }

        #endregion
    }
}
