using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core
{
    public class ItemFilter
    {
        public enum FilterMode
        {
            All         = 0,
            Conflict    = 1,
            NoConflict  = 2
        }

        FilterMode _mode;

        public FilterMode Mode
        {
            get 
            {
                return _mode;
            }
            set
            {
                _mode = value;
            }
    
        }

        private bool IsKnownEnum(XmlNode methodReturnNode)
        {
            string type = methodReturnNode.Attributes["Type"].InnerText;
            string isExternal = methodReturnNode.Attributes["IsExternal"].InnerText;
            string sourceKey = methodReturnNode.Attributes["SourceKey"].InnerText;
            if ("True" == isExternal)
            {
                // look with key
                XmlNode lookupNode = methodReturnNode.OwnerDocument.DocumentElement.SelectSingleNode("Exernals/External[@Key='" + sourceKey + "']");
                if (null != lookupNode)
                    return true;
            }
            else
            {
                XmlNode lookupNode = null;
                if (sourceKey != "")
                {
                    // look with key
                    lookupNode = methodReturnNode.OwnerDocument.DocumentElement.SelectSingleNode("Solution/Projects/Project[@Key='" + sourceKey + "']");
                    if (null != lookupNode)
                        return true;
                }
                else
                {  // look without key
                    lookupNode = methodReturnNode.OwnerDocument.DocumentElement.SelectSingleNode("Solution/Projects");
                    XmlNode targetType = lookupNode.SelectSingleNode("Project/Enums/Enum[@Name='" + type + "']");
                    if (null != targetType)
                        return true;
                }

              
            }

            return false;
        }
        
        private bool IsKnownScalarType(XmlNode methodReturnNode)
        {
            string type = methodReturnNode.Attributes["Type"].InnerText;

            switch (type)
            {
                case "int":
                case "Int16":
                case "Int32":
                case "Int64":
                case "double":
                case "Double":
                case "string":
                case "String":
                case "bool":
                case "float":
                    return true;
            }

            return false;
        }

        private bool IsKnownComProxyType(XmlNode methodReturnNode)
        {
            string type       = methodReturnNode.Attributes["Type"].InnerText;
            string isExternal = methodReturnNode.Attributes["IsExternal"].InnerText;
            string sourceKey  = methodReturnNode.Attributes["SourceKey"].InnerText;
            if ("True" == isExternal)
            {
                // look with key
                XmlNode lookupNode = methodReturnNode.OwnerDocument.DocumentElement.SelectSingleNode("Exernals/External[@Key='" + sourceKey + "']");
                if (null != lookupNode)
                    return true;
            }
            else
            {   
               
                XmlNode lookupNode = null;
                if (sourceKey != "")
                { 
                    // look with key
                    lookupNode = methodReturnNode.OwnerDocument.DocumentElement.SelectSingleNode("Solution/Projects/Project[@Key='" + sourceKey + "']");
                    if (null != lookupNode)
                        return true;
                }
                else
                { 
                    // look without key
                    lookupNode = methodReturnNode.OwnerDocument.DocumentElement.SelectSingleNode("Solution/Projects");
                   
                    XmlNode targetType = lookupNode.SelectSingleNode("Project/Interfaces/Interface[@Caption='" + type + "']");
                    if (null != targetType)
                        return true;

                    targetType = lookupNode.SelectSingleNode("Project/DispatchInterfaces/Interface[@Caption='" + type + "']");
                    if (null != targetType)
                        return true;

                    targetType = lookupNode.SelectSingleNode("Project/CoClasses/CoClass[@Caption='" + type + "']");
                    if (null != targetType)
                        return true;
                }
              
            }

            return false;
        }

        private bool MethodeNodeHasConflict(XmlNode methodNode)
        {
            #region Check Name

            string caption = methodNode.Attributes["Caption"].InnerText;
            if (caption.IndexOf(" ") > -1)
                return true;
            if (caption.IndexOfAny(System.IO.Path.GetInvalidPathChars()) > -1)
                return true;

            #endregion

            return false;
        }

        private bool ParametersNodeHasConflict(XmlNode parametersNode)
        {
            XmlNode returnNode = parametersNode.SelectSingleNode("ReturnValue");
            string returnValue = returnNode.Attributes["Type"].InnerText;
            if ("void" != returnValue)
            {
                string isComProxy = returnNode.Attributes["IsComProxy"].InnerText;
                if ("True" == isComProxy)
                {
                    bool returnTypeFound = IsKnownComProxyType(returnNode);
                    if (false == returnTypeFound)
                        return true;
                }
                else
                {
                    if ("TKIND_ENUM" == returnNode.Attributes["TypeKind"].InnerText)
                    {
                        bool returnEnumFound = IsKnownEnum(returnNode);
                        if (false == returnEnumFound)
                            return true;
                    }
                    else
                    {
                        bool returnScalarFound = IsKnownScalarType(returnNode);
                        if (false == returnScalarFound)
                            return true;
                    }
                }
            }

            return false;
        }

        private bool FilterMethodeNodePassed(XmlNode methodNode, ItemFilter.FilterMode mode)
        {
            bool hasConflict = MethodeNodeHasConflict(methodNode);
            switch (mode)
            {
                case FilterMode.Conflict:
                    return hasConflict;
                case FilterMode.NoConflict:
                    return !hasConflict;
            }
            return true;
        }

        private bool FilterParametersNodePassed(XmlNode parametersNode, ItemFilter.FilterMode mode)
        {
            bool hasConflict = ParametersNodeHasConflict(parametersNode);
            switch (mode)
            {
                case FilterMode.Conflict:
                    return hasConflict;
                case FilterMode.NoConflict:
                    return !hasConflict;
            }
            return true;
        }

        public bool FilterPassed(XmlNode node, ItemFilter.FilterMode mode)
        {
            if (mode == FilterMode.All)
                return true;

            string nodeType = node.Name;
            switch (nodeType)
            {
                case "Method":
                {
                    return FilterMethodeNodePassed(node, mode);
                }
                case "Parameters":
                {
                    return FilterParametersNodePassed(node, mode);
                }
            }

            throw (new ArgumentException("Unkown Node Type: " + node.Name));
        }

        public bool FilterPassed(XmlNode node)
        {
            return FilterPassed(node, _mode);
        }
    }
}
