﻿using System;
using System.Runtime.InteropServices; 
using System.Collections.Generic;
using System.Text;
using System.Xml;
using TLI;

namespace LateBindingApi.CodeGenerator.Core
{
    partial class CoClassHandler
    {
        private void AddProperty(XmlNode objectNode, TLI.MemberInfo methodInfo, string componentKey)
        {
            int optionalParameterCount = methodInfo.Parameters.OptionalCount;
            XmlNode methodsNode = objectNode.SelectSingleNode("Properties");
           
            XmlNode methodNode = LookForProperty(methodsNode, methodInfo, false);
            if (null == methodNode)
                methodNode = AddPropertyNodeToObject(methodsNode, methodInfo, componentKey, false);

            AddInvokeKindToPropertyNode(methodNode, methodInfo.InvokeKind);
            AddComponentKeyToPropertyNode(methodNode, componentKey);

            if (optionalParameterCount > 0)
            {
                methodNode = LookForProperty(methodsNode, methodInfo, true);
                if (null == methodNode)
                    AddPropertyNodeToObject(methodsNode, methodInfo, componentKey, true);
                else
                {
                    AddInvokeKindToPropertyNode(methodNode, methodInfo.InvokeKind);
                    AddComponentKeyToPropertyNode(methodNode, componentKey);
                }
            }
        }

        private void AddInvokeKindToPropertyNode(XmlNode methodNode, InvokeKinds kinds)
        {
            string kindString = kinds.ToString();
            methodNode.Attributes[kindString].InnerText = "true";
        }

        private void AddComponentKeyToPropertyNode(XmlNode methodNode, string componentKey)
        {
            XmlNode componentsNode = methodNode.SelectSingleNode("Components");
            XmlNode componentNode = componentsNode.SelectSingleNode(componentKey);
            if (null == componentNode)
            {
                componentNode = _parent.COMTree.CreateElement(componentKey);
                componentsNode.AppendChild(componentNode);
            }
        }
        
        private XmlNode AddPropertyNodeToObject(XmlNode methodsNode, TLI.MemberInfo methodInfo, string componentKey, bool withOptional)
        {
            XmlNode methodNode = _parent.COMTree.CreateElement("Property");
            methodsNode.AppendChild(methodNode);
            
            XmlAttribute attribute = _parent.COMTree.CreateAttribute("Name");
            attribute.InnerText = methodInfo.Name;
            methodNode.Attributes.Append(attribute);

            attribute = _parent.COMTree.CreateAttribute("IsExternalType");
            attribute.InnerText = methodInfo.ReturnType.IsExternalType.ToString();
            methodNode.Attributes.Append(attribute);

            XmlNode definitionNode = _parent.COMTree.CreateElement("Definition");
            methodNode.AppendChild(definitionNode);
            _parent.AddAtrributeToNode(definitionNode, "GUID", "");
            if (false == methodInfo.ReturnType.IsExternalType)
            {
                if (null != methodInfo.ReturnType.TypeInfo)
                    definitionNode.Attributes["GUID"].InnerText = XmlConvert.EncodeName(methodInfo.ReturnType.TypeInfo.Parent.GUID);
            }
            else
            {
                definitionNode.Attributes["GUID"].InnerText =XmlConvert.EncodeName( methodInfo.ReturnType.TypeLibInfoExternal.GUID);
            }

            attribute = _parent.COMTree.CreateAttribute("ReturnType");
            attribute.InnerText = _parent.GetReturnType(methodInfo);
            methodNode.Attributes.Append(attribute);

            attribute = _parent.COMTree.CreateAttribute("ReturnTypeKind");
            attribute.InnerText = "";
            methodNode.Attributes.Append(attribute);
            if (null != methodInfo.ReturnType.TypeInfo)
                attribute.InnerText = methodInfo.ReturnType.TypeInfo.TypeKind.ToString();  

            attribute = _parent.COMTree.CreateAttribute("IsRef");
            attribute.InnerText = "False";
            methodNode.Attributes.Append(attribute);

            attribute = _parent.COMTree.CreateAttribute("IsArray");
            attribute.InnerText = "False";
            methodNode.Attributes.Append(attribute);

            attribute = _parent.COMTree.CreateAttribute("INVOKE_PROPERTYGET");
            attribute.InnerText = "false";
            methodNode.Attributes.Append(attribute);

            attribute = _parent.COMTree.CreateAttribute("INVOKE_PROPERTYPUT");
            attribute.InnerText = "false";
            methodNode.Attributes.Append(attribute);

            attribute = _parent.COMTree.CreateAttribute("INVOKE_PROPERTYPUTREF");
            attribute.InnerText = "false";
            methodNode.Attributes.Append(attribute);
            methodNode.Attributes[methodInfo.InvokeKind.ToString()].InnerText = "true";

            XmlNode componentsNode = _parent.COMTree.CreateElement("Components");
            methodNode.AppendChild(componentsNode);
            AddComponentKeyToMethodNode(methodNode, componentKey);

            XmlNode paramsNode = _parent.COMTree.CreateElement("Parameters");
            methodNode.AppendChild(paramsNode);

            AddParametersToPropertyNode(methodNode, methodInfo, withOptional);

            return methodNode;
        }

        private void AddParametersToPropertyNode(XmlNode methodNode, TLI.MemberInfo methodInfo, bool withOptional)
        {
            XmlNode parametersNode = methodNode.SelectSingleNode("Parameters");
            foreach (ParameterInfo paramInfo in methodInfo.Parameters)
            {
                if (false == withOptional)
                {
                    if (paramInfo.Optional == true)
                        break;
                }

                XmlNode paramNode = parametersNode.OwnerDocument.CreateElement("Param");
                parametersNode.AppendChild(paramNode);

                XmlAttribute attribute = _parent.COMTree.CreateAttribute("Name");
                attribute.InnerText = paramInfo.Name;
                paramNode.Attributes.Append(attribute);

                attribute = _parent.COMTree.CreateAttribute("Type");
                attribute.InnerText = _parent.GetParameterType(paramInfo);
                paramNode.Attributes.Append(attribute);

                attribute = _parent.COMTree.CreateAttribute("ReturnTypeKind");
                attribute.InnerText = "";
                paramNode.Attributes.Append(attribute);
                if (null != paramInfo.VarTypeInfo.TypeInfo)
                    attribute.InnerText = paramInfo.VarTypeInfo.TypeInfo.TypeKind.ToString();  

                attribute = _parent.COMTree.CreateAttribute("Optional");
                attribute.InnerText = paramInfo.Optional.ToString();
                paramNode.Attributes.Append(attribute);

                attribute = _parent.COMTree.CreateAttribute("IsExternalType");
                attribute.InnerText = paramInfo.VarTypeInfo.IsExternalType.ToString();
                paramNode.Attributes.Append(attribute);

                XmlNode definitionNode = _parent.COMTree.CreateElement("Definition");
                paramNode.AppendChild(definitionNode);
                _parent.AddAtrributeToNode(definitionNode, "GUID", "");
                if (false == paramInfo.VarTypeInfo.IsExternalType)
                {
                    if (null != paramInfo.VarTypeInfo.TypeInfo)
                        definitionNode.Attributes["GUID"].InnerText = XmlConvert.EncodeName(paramInfo.VarTypeInfo.TypeInfo.Parent.GUID);
                }
                else
                {
                    definitionNode.Attributes["GUID"].InnerText = XmlConvert.EncodeName(paramInfo.VarTypeInfo.TypeLibInfoExternal.GUID);
                }

                attribute = _parent.COMTree.CreateAttribute("IsRef");
                attribute.InnerText = "False";
                paramNode.Attributes.Append(attribute);

                attribute = _parent.COMTree.CreateAttribute("IsArray");
                attribute.InnerText = "False";
                paramNode.Attributes.Append(attribute);
            }
        }

        private XmlNode LookForProperty(XmlNode methodsNode, TLI.MemberInfo methodInfo, bool withOptionals)
        {
            XmlNode returnNode = null;

            foreach (XmlNode nodeMethod in methodsNode.ChildNodes)
            {
                string methodName = methodInfo.Name;
                string methodName2 = nodeMethod.Attributes["Name"].InnerText;

                if (methodName == methodName2)
                {
                    XmlNode paramsNode = nodeMethod.SelectSingleNode("Parameters");
                    Parameters paramInfos = methodInfo.Parameters;

                    int methodInfoParametersCount = methodInfo.Parameters.Count;
                    if (withOptionals == false)
                        methodInfoParametersCount -= methodInfo.Parameters.OptionalCount;

                    if (paramsNode.ChildNodes.Count == methodInfoParametersCount)
                    {
                        for (int i = 1; i <= methodInfoParametersCount; i++)
                        {
                            ParameterInfo paramInfo = methodInfo.Parameters[(short)i];
                            if ((paramInfo.Optional == true) && (withOptionals == true))
                            {
                                Marshal.ReleaseComObject(paramInfo);
                                break;
                            }

                            #region check properties
                            XmlNode paramNode = paramsNode.ChildNodes[i - 1];
                            string paramType = paramInfo.VarTypeInfo.VarType.ToString();
                            string paramType2 = paramNode.Attributes["Type"].InnerText;
                            string paramName = paramInfo.Name;
                            string paramName2 = paramNode.Attributes["Name"].InnerText;

                            string paramOptional = paramInfo.Optional.ToString();
                            string paramOptional2 = paramNode.Attributes["Optional"].InnerText;

                            if ((paramType != paramType2) || (paramName != paramName2) || (paramOptional != paramOptional2))
                            {
                                Marshal.ReleaseComObject(paramInfo);
                                break;
                            }
                            #endregion

                            Marshal.ReleaseComObject(paramInfo);
                        }

                        Marshal.ReleaseComObject(paramInfos);
                        returnNode = nodeMethod;
                        return nodeMethod;
                    }

                    Marshal.ReleaseComObject(paramInfos);
                }

            }

            return returnNode;
        }

        private bool IsClassProperty(TLI.MemberInfo memberInfo)
        {
            if (memberInfo.DescKind == DescKinds.DESCKIND_FUNCDESC)
            {
                if ((memberInfo.InvokeKind == InvokeKinds.INVOKE_PROPERTYGET) ||
                    (memberInfo.InvokeKind == InvokeKinds.INVOKE_PROPERTYPUT) ||
                    (memberInfo.InvokeKind == InvokeKinds.INVOKE_PROPERTYPUTREF))
                {
                     return true;
                }
                else
                    return false;
            }
            else
            {
                    return false;
            }

        }
    }
}
