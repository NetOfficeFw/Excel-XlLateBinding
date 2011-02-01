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
    internal class MethodHandler
    {
        #region Fields

        COMComponentReader _parent;
        
        #endregion

        #region Construction

        internal MethodHandler(COMComponentReader parent)
        {
            _parent = parent;
        }

        #endregion

        #region Methods

        internal void AddMethod(XmlNode componentNode, XmlNode projectNode, XmlNode objectNode, TLI.MemberInfo methodInfo)
        {
            int optionalParameterCount = methodInfo.Parameters.OptionalCount;
            XmlNode methodsNode = objectNode.SelectSingleNode("Methods");

            XmlNode methodNode = GetMethodNode(methodsNode, methodInfo);
            if (null == methodNode)
                methodNode = AddNewMethodNodeToObject(componentNode, methodsNode, methodInfo);

            AddDispIdMethodNode(componentNode, methodNode, methodInfo);
            AddComponentRefToMethodNode(componentNode, methodNode);

            AddParametersToMethodNode(componentNode, methodNode, methodInfo, false);
            AddParametersToMethodNode(componentNode, methodNode, methodInfo, true);
        }

        private bool MethodHasReturnValue(TLI.MemberInfo methodInfo)
        {
            switch (methodInfo.ReturnType.VarType)
            {
                case TliVarType.VT_VOID:
                    return false;
                default:
                    return true;
            }
        }

        private void AddDispIdMethodNode(XmlNode componentNode, XmlNode methodNode, TLI.MemberInfo methodInfo)
        {
            string key = componentNode.Attributes["Key"].InnerText;
            XmlNode refDipsIdsNode = methodNode.SelectSingleNode("DispIds");
            XmlNode refIdNode = refDipsIdsNode.SelectSingleNode("DispId[@ComponentKey = '" + key + "']");
            if (null == refIdNode)
            {
                refIdNode = _parent.Document.CreateElement("DispId");
                Utils.AddAtrributeToNode(refIdNode, "ComponentKey", key);
                Utils.AddAtrributeToNode(refIdNode, "Id", methodInfo.MemberId.ToString());
                refDipsIdsNode.AppendChild(refIdNode);
            }
        }

        private void AddComponentRefToMethodNode(XmlNode componentNode, XmlNode methodNode)
        {
            string key = componentNode.Attributes["Key"].InnerText;
            XmlNode refComponentsNode = methodNode.SelectSingleNode("RefComponents");
            XmlNode refComponentNode = refComponentsNode.SelectSingleNode("RefComponent[@Key = '" + key + "']");
            if (null == refComponentNode)
            {
                refComponentNode = _parent.Document.CreateElement("RefComponent");
                Utils.AddAtrributeToNode(refComponentNode, "Key", key);
                refComponentsNode.AppendChild(refComponentNode);
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

        private XmlNode AddNewMethodNodeToObject(XmlNode componentNode, XmlNode methodsNode, TLI.MemberInfo methodInfo)
        {
            XmlNode methodNode = _parent.Document.CreateElement("Method");
            methodsNode.AppendChild(methodNode);

            Utils.AddAtrributeToNode(methodNode, "Name", methodInfo.Name);
            Utils.AddAtrributeToNode(methodNode, "Caption", methodInfo.Name);
            Utils.AddAtrributeToNode(methodNode, "Key", XmlConvert.EncodeName(Guid.NewGuid().ToString()));
            Utils.AddAtrributeToNode(methodNode, "HasReturnType", MethodHasReturnValue(methodInfo).ToString());
            
            Utils.AddChildToNode(methodNode, "DispIds");
            Utils.AddChildToNode(methodNode, "RefComponents");
           
            return methodNode;
        }

        private XmlNode GetMethodNode(XmlNode methodsNode, TLI.MemberInfo methodInfo)
        {
            return methodsNode.SelectSingleNode("Method[@Name ='" + methodInfo.Name + "']");
        }

        private XmlNode GetParametersNode(XmlNode methodNode, TLI.MemberInfo propertyInfo, bool withOptionals)
        {
            int targetParamsCount = propertyInfo.Parameters.Count;
            if (false == withOptionals)
                targetParamsCount -= propertyInfo.Parameters.OptionalCount;

            List<XmlNode> listNodes = GetParametersList(methodNode, propertyInfo, targetParamsCount);
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

                XmlNode x = parametersNode.SelectSingleNode("Parameter[@Name='" + paramInfo.Name + "'" + " and " + "@Type='" + _parent.GetParameterType(paramInfo) + "']");
                if (null != x)
                {
                    Marshal.ReleaseComObject(paramInfo);
                    return parametersNode;
                }

                Marshal.ReleaseComObject(paramInfo);
            }

            return null;
        }
       
        private void AddParametersToMethodNode(XmlNode componentNode, XmlNode methodNode, TLI.MemberInfo methodInfo, bool withOptionals)
        {
            XmlNode parametersNode = GetParametersNode(methodNode, methodInfo, withOptionals);
            if (null == parametersNode)
            {
                parametersNode = _parent.Document.CreateElement("Parameters");
                Utils.AddChildToNode(parametersNode, "RefComponents");
                Utils.AddChildToNode(parametersNode, "ReturnValue");
                methodNode.AppendChild(parametersNode);

                foreach (ParameterInfo paramInfo in methodInfo.Parameters)
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
            AddReturnValueToParametersNode(parametersNode, methodInfo);
            AddComponentRefToParametersNode(componentNode, parametersNode);
        }

        #endregion
    }
}
