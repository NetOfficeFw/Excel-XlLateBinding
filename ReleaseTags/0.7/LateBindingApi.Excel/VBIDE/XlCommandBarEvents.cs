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
    public class XlCommandBarEvents : XlNonCreatable
    {   
        #region Construction

        internal XlCommandBarEvents(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {

        }

        #endregion
    }
}
