using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Schema;

using TLI;

namespace LateBindingApi.CodeGenerator.Core
{
    internal class PropertyHandler
    {
        #region Fields

        COMComponentReader _parent;

        #endregion

        #region Construction

        internal PropertyHandler(COMComponentReader parent)
        {
            _parent = parent;
        }

        #endregion

        #region Methods

        internal void AddProperty(XmlNode componentNode, XmlNode projectNode, XmlNode objectNode, TLI.MemberInfo propertyInfo)
        {   
            int optionalParameterCount = propertyInfo.Parameters.OptionalCount;
            XmlNode propertiesNode = objectNode.SelectSingleNode("Properties");

            XmlNode propertyNode = GetPropertyNode(propertiesNode, propertyInfo);
            if (null == propertyNode)
                propertyNode = AddNewPropertyNodeToObject(componentNode, propertiesNode, propertyInfo);

            AddInvokeKindToPropertyNode(propertyNode, propertyInfo);
            AddDispIdPropertyNode(componentNode, propertyNode, propertyInfo);
            AddComponentRefToPropertyNode(componentNode, propertyNode);
            
            AddParametersToPropertyNode(componentNode, propertyNode, propertyInfo, false);
            AddParametersToPropertyNode(componentNode, propertyNode, propertyInfo, true);

        }

        private void AddDispIdPropertyNode(XmlNode componentNode, XmlNode propertyNode, TLI.MemberInfo propertyInfo)
        {
            string key = componentNode.Attributes["Key"].InnerText;
            XmlNode refDipsIdsNode = propertyNode.SelectSingleNode("DispIds");
            XmlNode refIdNode = refDipsIdsNode.SelectSingleNode("DispId[@ComponentKey = '" + key + "']");
            if (null == refIdNode)
            {
                refIdNode = _parent.Document.CreateElement("DispId");
                Utils.AddAtrributeToNode(refIdNode, "ComponentKey", key);
                Utils.AddAtrributeToNode(refIdNode, "Id", propertyInfo.MemberId.ToString());
                refDipsIdsNode.AppendChild(refIdNode);
            }
        }
       
        private void AddComponentRefToPropertyNode(XmlNode componentNode, XmlNode propertyNode)
        {
            string key = componentNode.Attributes["Key"].InnerText;
            XmlNode refComponentsNode = propertyNode.SelectSingleNode("RefComponents");
            XmlNode refComponentNode = refComponentsNode.SelectSingleNode("RefComponent[@Key = '" + key + "']");
            if (null == refComponentNode)
            {
                refComponentNode = _parent.Document.CreateElement("RefComponent");
                Utils.AddAtrributeToNode(refComponentNode, "Key", key);
                refComponentsNode.AppendChild(refComponentNode);
            }
        }

        private void AddComponentRefToParametersNode(XmlNode componentNode, XmlNode parametersNode)
        {
            string key = componentNode.Attributes["Key"].InnerText;
            XmlNode refComponentsNode = parametersNode.SelectSingleNode("RefComponents");
            XmlNode refComponentNode = refComponentsNode.SelectSingleNode("RefComponent[@Key = '" + key + "']");
            if (null == refComponentNode)
            {
                refComponentNode = _parent.Document.CreateElement("RefComponent");
                Utils.AddAtrributeToNode(refComponentNode, "Key", key);
                refComponentsNode.AppendChild(refComponentNode);
            }
        }

        private void AddInvokeKindToPropertyNode(XmlNode propertyNode, TLI.MemberInfo propertyInfo)
        {
            InvokeKinds kinds = propertyInfo.InvokeKind;

            string kindExists = propertyNode.Attributes["InvokeKind"].InnerText;
            switch (kinds)
            {
                case InvokeKinds.INVOKE_PROPERTYGET:
                    if ((kindExists != "INVOKE_PROPERTYPUT") && (kindExists != "INVOKE_PROPERTYPUTREF"))
                        propertyNode.Attributes["InvokeKind"].InnerText = kinds.ToString();
                    break;
                case InvokeKinds.INVOKE_PROPERTYPUT:
                case InvokeKinds.INVOKE_PROPERTYPUTREF:
                    propertyNode.Attributes["InvokeKind"].InnerText = kinds.ToString();
                    break;

            }
        }
     
        private void AddReturnValueToParametersNode(XmlNode parametersNode, TLI.MemberInfo methodInfo)
        {
            XmlNode methodReturnNode = parametersNode.SelectSingleNode("ReturnValue");

            Utils.AddAtrributeToNode(methodReturnNode, "IsComProxy", _parent.IsComProxy(methodInfo.ReturnType).ToString());
            Utils.AddAtrributeToNode(methodReturnNode, "IsExternal", methodInfo.ReturnType.IsExternalType.ToString());
            Utils.AddAtrributeToNode(methodReturnNode, "IsArray", "False");
            Utils.AddAtrributeToNode(methodReturnNode, "Type", _parent.GetReturnType(methodInfo));

            string typeKind = "";
            if (null != methodInfo.ReturnType.TypeInfo)
                typeKind = methodInfo.ReturnType.TypeInfo.TypeKind.ToString();
            Utils.AddAtrributeToNode(methodReturnNode, "TypeKind", typeKind);

            Utils.AddAtrributeToNode(methodReturnNode, "SourceKey", _parent.GetSourceKey(methodInfo.ReturnType));
        }

        private XmlNode AddNewPropertyNodeToObject(XmlNode componentNode,  XmlNode propertiesNode, TLI.MemberInfo propertyInfo)
        {
            XmlNode propertyNode = _parent.Document.CreateElement("Property");
            propertiesNode.AppendChild(propertyNode);

            Utils.AddChildToNode(propertyNode, "DispIds");
            Utils.AddChildToNode(propertyNode, "RefComponents");

            Utils.AddAtrributeToNode(propertyNode, "Name", propertyInfo.Name);
            Utils.AddAtrributeToNode(propertyNode, "Caption", propertyInfo.Name);
            Utils.AddAtrributeToNode(propertyNode, "Key", XmlConvert.EncodeName(Guid.NewGuid().ToString()));
            Utils.AddAtrributeToNode(propertyNode, "InvokeKind", propertyInfo.InvokeKind.ToString());
            Utils.AddAtrributeToNode(propertyNode, "SourceKey", _parent.GetSourceKey(propertyInfo.ReturnType));

            string typeKind = "";
            if (null != propertyInfo.ReturnType.TypeInfo)
                typeKind = propertyInfo.ReturnType.TypeInfo.TypeKind.ToString();
            Utils.AddAtrributeToNode(propertyNode, "TypeKind", typeKind);
  
            return propertyNode;
        }

        private XmlNode GetPropertyNode(XmlNode propertiesNode, TLI.MemberInfo propertyInfo)
        {
            return propertiesNode.SelectSingleNode("Property[@Name ='" + propertyInfo.Name + "']");
        }
        
        private List<XmlNode> GetParametersList(XmlNode methodNode, TLI.MemberInfo propertyInfo, int targetParamsCount)
        {

            List<XmlNode> returnList = new List<XmlNode>();
            XmlNodeList paramsNodes = methodNode.SelectNodes("Parameters");
            foreach (XmlNode paramNode in paramsNodes)
            {
                XmlNodeList paramNodes = paramNode.SelectNodes("Parameter");
                if (paramNodes.Count == targetParamsCount)
                {
                    returnList.Add(paramNode);
                }
            }

            return returnList;
        }

        private XmlNode GetParametersNode(XmlNode propertyNode, TLI.MemberInfo propertyInfo, bool withOptionals)
        {
            int targetParamsCount = propertyInfo.Parameters.Count;
            if (false == withOptionals)
                targetParamsCount -= propertyInfo.Parameters.OptionalCount;

            List<XmlNode> listNodes = GetParametersList(propertyNode, propertyInfo, targetParamsCount);
            if ((targetParamsCount == 0) && (listNodes.Count == 1))
                return listNodes[0];

            if (listNodes.Count == 0)
                return null;

            for (int i = 0; i < propertyInfo.Parameters.Count; i++)
            {
                ParameterInfo paramInfo = propertyInfo.Parameters[(short)(i + 1)];
                XmlNode parametersNode = listNodes[i];
                if ((withOptionals == false) && (paramInfo.Optional == true))
                {
                    Marshal.ReleaseComObject(paramInfo);
                    return null;
                }

                XmlNode x = parametersNode.SelectSingleNode("Parameter[@Name='" + paramInfo.Name + "'" + " and " +
                                                    "@Type='" + _parent.GetParameterType(paramInfo) + "']");
                if (null != x)
                {
                    Marshal.ReleaseComObject(paramInfo);
                    return parametersNode;
                }

                Marshal.ReleaseComObject(paramInfo);
            }

            return null;
        }

        private void AddParametersToPropertyNode(XmlNode componentNode, XmlNode propertyNode, TLI.MemberInfo propertyInfo, bool withOptionals)
        {
            if (((propertyInfo.Parameters.Count - propertyInfo.Parameters.OptionalCount) == 0) && (false == withOptionals))
                return;

            XmlNode parametersNode = GetParametersNode(propertyNode, propertyInfo, withOptionals);
            if (null == parametersNode)
            {
                parametersNode = _parent.Document.CreateElement("Parameters");
                Utils.AddChildToNode(parametersNode, "ReturnValue");
                Utils.AddChildToNode(parametersNode, "RefComponents");
                propertyNode.AppendChild(parametersNode);

                foreach (ParameterInfo paramInfo in propertyInfo.Parameters)
                {
                    if ((false == withOptionals) && (paramInfo.Optional == true))
                        break;

                    XmlNode paramNode = parametersNode.OwnerDocument.CreateElement("Parameter");
                    parametersNode.AppendChild(paramNode);

                    Utils.AddAtrributeToNode(paramNode, "Name", paramInfo.Name);
                    Utils.AddAtrributeToNode(paramNode, "Caption", paramInfo.Name);
                    Utils.AddAtrributeToNode(paramNode, "IsExternal", paramInfo.VarTypeInfo.IsExternalType.ToString());
                    Utils.AddAtrributeToNode(paramNode, "IsComProxy", _parent.IsComProxy(paramInfo.VarTypeInfo).ToString());
                    Utils.AddAtrributeToNode(paramNode, "IsOptional", paramInfo.Optional.ToString());
                    Utils.AddAtrributeToNode(paramNode, "IsRef", "false");
                    Utils.AddAtrributeToNode(paramNode, "IsArray", "false");
                    Utils.AddAtrributeToNode(paramNode, "Type", _parent.GetParameterType(paramInfo));

                    string typeKind = "";
                    if (null != paramInfo.VarTypeInfo.TypeInfo)
                        typeKind = paramInfo.VarTypeInfo.TypeInfo.TypeKind.ToString();
                    Utils.AddAtrributeToNode(paramNode, "TypeKind", typeKind);

                    Utils.AddAtrributeToNode(paramNode, "SourceKey", _parent.GetSourceKey(paramInfo.VarTypeInfo));
                }
            }
            AddReturnValueToParametersNode(parametersNode, propertyInfo);
            AddComponentRefToParametersNode(componentNode, parametersNode);
        }

        #endregion
    }
}
