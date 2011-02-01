using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Language
{
    /// <summary>
    /// Represents LanguageSettings Object
    /// </summary>
    public class XlLanguageSettings : XlNonCreatable
    {
        #region Construction

        internal XlLanguageSettings(IXlObject parentReference, object comReference): base(parentReference, comReference)
        {
        }

        #endregion

        #region Scalar Properties

        /// <summary>
        /// LanguageID 
        /// </summary>
        public int LanguageID(XlMsoAppLanguageID ID)
        {
            object returnValue  = InstanceType.InvokeMember("LanguageID", BindingFlags.GetProperty, null, ComReference, new object[1] { ID }, XlLateBindingApiSettings.XlThreadCulture);
            return (int)returnValue;    
        }

        #endregion  
    }
}
