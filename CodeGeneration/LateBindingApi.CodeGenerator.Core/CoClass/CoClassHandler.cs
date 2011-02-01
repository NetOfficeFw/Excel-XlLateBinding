using System.Xml;
using TLI;
using System.Runtime.InteropServices;

namespace LateBindingApi.CodeGenerator.Core
{
    internal partial class CoClassHandler
    {
        #region Fields

        COMComponentReader _parent;

        #endregion

        #region Construction
        
        internal CoClassHandler(COMComponentReader parent)        
        {
            _parent = parent;
        }

        #endregion

        #region Methods

        internal void ProceedClass(string componentKey, XmlNode classesNode, CoClassInfo itemClass)
        {
            string className = itemClass.Name;
            XmlNode classNode = LookForClassNode(classesNode, className);
            if (classNode == null)
                classNode = CreateClassNode(itemClass, classesNode, className);

            AddComponentKeyToClassNode(classNode, componentKey);
            AddClassInfo(componentKey, classNode, itemClass);
        }

        internal void AddComponentKeyToClassNode(XmlNode classNode, string componentKey)
        {
            XmlNode componentsNode = classNode.SelectSingleNode("Components");
            XmlNode componentNode = componentsNode.SelectSingleNode(componentKey);
            if (null == componentNode)
            {

                componentNode = _parent.COMTree.CreateElement(componentKey);
                componentsNode.AppendChild(componentNode);
            }
        }

        private XmlNode CreateClassNode(TLI.CoClassInfo itemClass, XmlNode classesNode, string className)
        {
            XmlNode classNode = _parent.COMTree.CreateElement("CoClass");
            XmlAttribute attribName = _parent.COMTree.CreateAttribute("Name");
            attribName.InnerText = className;
            classNode.Attributes.Append(attribName);

            attribName = _parent.COMTree.CreateAttribute("GUID");
            attribName.InnerText = XmlConvert.EncodeName(itemClass.GUID);
            classNode.Attributes.Append(attribName);            

            classNode.AppendChild(_parent.COMTree.CreateElement("Components"));
            classNode.AppendChild(_parent.COMTree.CreateElement("Properties"));
            classNode.AppendChild(_parent.COMTree.CreateElement("Methods"));
            XmlNode interfacesNode = _parent.COMTree.CreateElement("Interfaces");
            classNode.AppendChild(interfacesNode);

            XmlNode faceNode = _parent.COMTree.CreateElement("Implied");
            interfacesNode.AppendChild(faceNode);
            foreach (InterfaceInfo faceInfo in itemClass.Interfaces)
            {
                XmlNode newFaceNode = _parent.COMTree.CreateElement("Interface");
                XmlAttribute attribFace = _parent.COMTree.CreateAttribute("Name");
                attribFace.InnerText = faceInfo.Name;
                newFaceNode.Attributes.Append(attribFace);

                faceNode.AppendChild(newFaceNode);
            }

            faceNode = _parent.COMTree.CreateElement("VTable");
            interfacesNode.AppendChild(faceNode);
            if (null != itemClass.DefaultInterface.VTableInterface)
            {
                foreach (InterfaceInfo faceInfo in itemClass.DefaultInterface.VTableInterface.ImpliedInterfaces)
                {
                    XmlNode newFaceNode = _parent.COMTree.CreateElement("Interface");
                    XmlAttribute attribFace = _parent.COMTree.CreateAttribute("Name");
                    attribFace.InnerText = faceInfo.Name;
                    newFaceNode.Attributes.Append(attribFace);

                    faceNode.AppendChild(newFaceNode);
                }
            }

            faceNode = _parent.COMTree.CreateElement("Events");
            interfacesNode.AppendChild(faceNode);
            if (null != itemClass.DefaultEventInterface)
            {
                XmlNode newFaceNode = _parent.COMTree.CreateElement("Interface");
                XmlAttribute attribFace = _parent.COMTree.CreateAttribute("Name");
                string eventInterfaceName = itemClass.DefaultEventInterface.Name;
                attribFace.InnerText = eventInterfaceName;
                newFaceNode.Attributes.Append(attribFace);

                faceNode.AppendChild(newFaceNode);
            }

            classesNode.AppendChild(classNode);
            return classNode;
        }

        private XmlNode LookForClassNode(XmlNode classesNode, string className)
        {
            foreach (XmlNode classNode in classesNode.ChildNodes)
            {
                if (className == classNode.Attributes["Name"].InnerText)
                    return classNode;
            }
            return null;
        }

        private void AddClassInfo(string componentKey, XmlNode classNode, CoClassInfo classInfo)
        {
            InterfaceInfo defaultInterface = classInfo.DefaultInterface;
            InterfaceInfo defaultEventInterface = classInfo.DefaultEventInterface;
            foreach (TLI.MemberInfo itemMember in defaultInterface.Members)
            {
                if (true == IsClassMethod(itemMember))
                {
                    AddMethod(classNode, itemMember, componentKey);
                }
                else if (true == IsClassProperty(itemMember))
                {
                    AddProperty(classNode, itemMember, componentKey);
                }
                Marshal.ReleaseComObject(itemMember);
            }

            Marshal.ReleaseComObject(defaultInterface);
        }

        #endregion
    }
}
