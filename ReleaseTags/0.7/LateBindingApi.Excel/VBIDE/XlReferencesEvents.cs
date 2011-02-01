using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel.VBIDE
{
    public class XlReferencesEvents : XlNonCreatable
    {  
        #region Construction

        internal XlReferencesEvents(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion
    }
}
