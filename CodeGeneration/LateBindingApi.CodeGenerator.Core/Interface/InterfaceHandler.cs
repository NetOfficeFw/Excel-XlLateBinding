using System.Xml;
using TLI;
using System.Runtime.InteropServices;

namespace LateBindingApi.CodeGenerator.Core
{
    internal partial class InterfaceHandler
    { 
        #region Fields

        COMComponentReader _parent;

        #endregion

        #region Construction

        internal InterfaceHandler(COMComponentReader parent)        
        {
            _parent = parent;
        }

        #endregion

        #region Methods

        internal void ProceedInterface(string componentKey, XmlNode interfacesNode, InterfaceInfo itemInterface)
        {
            string interfaceName = itemInterface.Name;
            XmlNode faceNode = LookForInterfaceNode(interfacesNode, interfaceName);
            if (faceNode == null)
                faceNode = CreateInterfaceNode(itemInterface, interfacesNode, interfaceName);

            AddComponentKeyToInterfaceNode(faceNode, componentKey);
            AddInterfaceInfo(componentKey, faceNode, itemInterface);
        }

        private void AddComponentKeyToInterfaceNode(XmlNode faceNode, string componentKey)
        {
            XmlNode componentsNode = faceNode.SelectSingleNode("Components");
            XmlNode componentNode = componentsNode.SelectSingleNode(componentKey);
            if (null == componentNode)
            {

                componentNode = _parent.COMTree.CreateElement(componentKey);
                componentsNode.AppendChild(componentNode);
            }
        }

        private XmlNode LookForInterfaceNode(XmlNode interfacesNode, string interfaceName)
        {
            foreach (XmlNode interfaceNode in interfacesNode.ChildNodes)
            {
                if (interfaceName == interfaceNode.Attributes["Name"].InnerText)
                    return interfaceNode;
            }
            return null;
        }

        private void AddInterfaceInfo(string componentKey, XmlNode interfaceNode, InterfaceInfo interfaceInfo)
        {
            foreach (TLI.MemberInfo itemMember in interfaceInfo.Members)
            {
                if (true == IsInterfaceMethod(itemMember))
                {
                    AddMethod(interfaceNode, itemMember, componentKey);
                }
                else if (true == IsInterfaceProperty(itemMember))
                {
                    AddProperty(interfaceNode, itemMember, componentKey);
                }
                Marshal.ReleaseComObject(itemMember);
            }
        }

        internal XmlNode CreateInterfaceNode(TLI.InterfaceInfo itemInterface, XmlNode facesNode, string interfaceName)
        {
            XmlNode ifaceNode = _parent.COMTree.CreateElement("Interface");
            XmlAttribute attribName = _parent.COMTree.CreateAttribute("Name");
            attribName.InnerText = interfaceName;
            ifaceNode.Attributes.Append(attribName);

            attribName = _parent.COMTree.CreateAttribute("GUID");
            attribName.InnerText =  XmlConvert.EncodeName(itemInterface.GUID);
            ifaceNode.Attributes.Append(attribName);

            ifaceNode.AppendChild(_parent.COMTree.CreateElement("Components"));
            ifaceNode.AppendChild(_parent.COMTree.CreateElement("Properties"));
            ifaceNode.AppendChild(_parent.COMTree.CreateElement("Methods"));
            ifaceNode.AppendChild(_parent.COMTree.CreateElement("Events"));
            XmlNode interfacesNode = _parent.COMTree.CreateElement("Interfaces");
            ifaceNode.AppendChild(interfacesNode);

            XmlNode faceNode = _parent.COMTree.CreateElement("Implied");
            interfacesNode.AppendChild(faceNode);
            foreach (InterfaceInfo faceInfo in itemInterface.ImpliedInterfaces)
            {
                XmlNode newFaceNode = _parent.COMTree.CreateElement("Interface");
                XmlAttribute attribFace = _parent.COMTree.CreateAttribute("Name");
                attribFace.InnerText = faceInfo.Name;
                newFaceNode.Attributes.Append(attribFace);

                attribFace = _parent.COMTree.CreateAttribute("Type");
                attribFace.InnerText = "Implied";
                newFaceNode.Attributes.Append(attribFace);

                faceNode.AppendChild(newFaceNode);
            }

            faceNode = _parent.COMTree.CreateElement("VTable");
            interfacesNode.AppendChild(faceNode);
            if (null != itemInterface.VTableInterface)
            {
                foreach (InterfaceInfo faceInfo in itemInterface.VTableInterface.ImpliedInterfaces)
                {
                    XmlNode newFaceNode = _parent.COMTree.CreateElement("Interface");
                    XmlAttribute attribFace = _parent.COMTree.CreateAttribute("Name");
                    attribFace.InnerText = faceInfo.Name;
                    newFaceNode.Attributes.Append(attribFace);

                    attribFace = _parent.COMTree.CreateAttribute("Type");
                    attribFace.InnerText = "VTable";
                    newFaceNode.Attributes.Append(attribFace);

                    faceNode.AppendChild(newFaceNode);
                }
            }

            facesNode.AppendChild(ifaceNode);
            return ifaceNode;
        }

        #endregion
    }
}
