using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using TLI;

namespace LateBindingApi.CodeGenerator.Core
{
    class EnumHandler
    {
        #region Fields

        COMComponentReader _parent;
        
        #endregion
        
        #region Construction

        internal EnumHandler(COMComponentReader parent)
        {
            _parent = parent;
        }

        #endregion

        #region Methods

        internal XmlNode ProceedEnum(string componentKey, XmlNode enumsNode, TLI.ConstantInfo itemConstant)
        {
           
            XmlNode node = _parent.COMTree.CreateElement(itemConstant.Name);
            XmlNode enumNode = _parent.GetChildNodeByAttribute(enumsNode, "Name", itemConstant.Name);
            if (enumNode == null)
            {
                enumNode = _parent.COMTree.CreateElement("Enum");
                _parent.AddAtrributeToNode(enumNode, "Name", itemConstant.Name); 
                enumNode.AppendChild(_parent.COMTree.CreateElement("Components"));
                enumNode.AppendChild(_parent.COMTree.CreateElement("Members"));
                enumsNode.AppendChild(enumNode);

                AddEnumInfo(componentKey, enumNode, itemConstant);
            }
            else
            {
                AddEnumInfo(componentKey, enumNode, itemConstant);
            }

            return enumNode;
        }

        private void AddEnumInfo(string componentKey, XmlNode enumNode, ConstantInfo constInfo)
        {
            XmlNode componentsNode = enumNode.SelectSingleNode("Components");
            XmlNode componentNode = _parent.GetChildNodeByInnerText(componentsNode, componentKey);
            if (null == componentNode)
            {
                componentNode = _parent.COMTree.CreateElement("Component");
                componentNode.InnerText = componentKey; 
                componentsNode.AppendChild(componentNode);
            }
         
            XmlNode membersNode = enumNode.SelectSingleNode("Members");
            foreach (TLI.MemberInfo itemMember in constInfo.Members)
            {
                XmlNode enumMemberNode = _parent.GetChildNodeByAttribute(membersNode, "Name", itemMember.Name);
                if (null == enumMemberNode)
                {
                    enumMemberNode = _parent.COMTree.CreateElement("EnumMember");
                    _parent.AddAtrributeToNode(enumMemberNode, "Name", itemMember.Name);
                    _parent.AddAtrributeToNode(enumMemberNode, "Value", itemMember.Value.ToString());
                    membersNode.AppendChild(enumMemberNode);
                    _parent.AddNewChildNode(enumMemberNode, "Components");
                }
                AddComponentToEnumMember(componentKey, enumMemberNode);
            }
        }

        private void AddComponentToEnumMember(string componentKey, XmlNode enumMemberNode)
        {
            XmlNode memberComponentsNode = enumMemberNode.SelectSingleNode("Components");
            XmlNode componentNode = _parent.GetChildNodeByInnerText(memberComponentsNode, componentKey);
            if (null == componentNode)
            {
                componentNode = _parent.COMTree.CreateElement("Component");
                componentNode.InnerText = componentKey; 

                memberComponentsNode.AppendChild(componentNode);
            }           
        }
        
        #endregion
    }
}
