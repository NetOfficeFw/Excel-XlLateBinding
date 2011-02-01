﻿//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Collections.Generic;
using LateBindingApi.Core;

namespace LateBindingApi.Office.Utils
{
    public class FactoryInfo : IFactoryInfo
    {
        #region Field

        private string _prefix = "";
        private string _namespace = "LateBindingApi.Office";
        private Guid _componentGuid = new Guid("2DF8D04C-5BFA-101B-BDE5-00AA0044DE52");
        private Assembly _assembly;

        #endregion

        #region Construction

        public FactoryInfo()
        {
            _assembly = System.Reflection.Assembly.GetExecutingAssembly();
        }

        #endregion

        #region IFactoryInfo Members

        public string Prefix
        {
            get
            {
                return _prefix;
            }
        }

        public string Namespace
        {
            get
            {
                return _namespace;
            }
        }

        public Guid ComponentGuid
        {
            get
            {
                return _componentGuid;
            }
        }

        public Assembly Assembly
        {
            get
            {
                return _assembly;
            }
        }

        #endregion
    }
}
