using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.CodeGenerator.Core
{
    public class ElementCatalog
    {
        #region Fields

        private ElementEnums _enums;
        
        #endregion

        #region Properties

        public ElementEnums Enums
        {
            get 
            {
                return _enums;
            }
        }
        
        #endregion

        #region Construction

        public ElementCatalog()
        {
            _enums = new ElementEnums();
        }
        
        #endregion
        
        #region Methods
        
        internal void Clear()
        {
            _enums.Clear();
        }

        #endregion
    }
}
