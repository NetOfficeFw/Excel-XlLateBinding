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
    public partial class COMComponentReader
    {
        #region Fields
        
        TLIApplication           _typeLibApplication;
        XmlDocument              _document;
        MethodHandler            _methodHandler;
        PropertyHandler         _propertyHandler;

        #endregion

        #region Properties

        public XmlDocument Document
        {
            get 
            {
                return _document;
            }
        }
      
        #endregion

        #region Construction

        public COMComponentReader()
        {
            ResetDocument();
            _methodHandler = new MethodHandler(this);
            _propertyHandler = new PropertyHandler(this);
        }

        #endregion

        #region LoadSave Methods
        
        public void LoadCOMComponents(COMComponentReaderSettings settings)
        {
            _typeLibApplication = new TLIApplication();

            if (true == settings.CreateNewProject)
                ResetDocument();

            LoadExternalsComponents(settings);
            LoadComponents(settings);
            LoadSolution(settings);
            LoadEnums(settings);
            LoadInterfaces(settings, false);
            LoadInterfaces(settings, true);          
            LoadCoClasses(settings);

            Marshal.ReleaseComObject(_typeLibApplication);
            _typeLibApplication = null;
        }

        public void LoadProject(string fileName)
        {
            _document.Load(fileName);
            Utils.Validate(_document);
        }

        public void SaveProject(string fileName) 
        {
            
            string nameOfFile = System.IO.Path.GetFileNameWithoutExtension(fileName);
            string projectDirectory = System.IO.Path.Combine(Environment.CurrentDirectory, "ProjectDirectory" + nameOfFile);
            XmlNode componentsNode = _document.SelectSingleNode(XPathConstants.Components);
            
            if (false == System.IO.Directory.Exists(projectDirectory))
                System.IO.Directory.CreateDirectory(projectDirectory);

            foreach (XmlNode itemNode in componentsNode.ChildNodes)
            {
                string helpFile = itemNode.Attributes["HelpFile"].InnerText;
                helpFile = helpFile.Replace("\0", "");

                string helpFileName = System.IO.Path.GetFileName(helpFile);

                string containingFile = itemNode.Attributes["ContainingFile"].InnerText;
                string oldFileName = System.IO.Path.GetFileName(containingFile);
                if (oldFileName.IndexOf("}", 0, StringComparison.InvariantCultureIgnoreCase) > 0)
                {
                    oldFileName = oldFileName.Substring(oldFileName.IndexOf("}", 0, StringComparison.InvariantCultureIgnoreCase) + 1);
                }

                string binaryName = projectDirectory + "\\{" + itemNode.Attributes["Key"].InnerText + "}" + oldFileName;
                string helpName = projectDirectory + "\\{" + itemNode.Attributes["Key"].InnerText + "}" + helpFileName;

                if (false == System.IO.File.Exists(containingFile))
                    containingFile = System.IO.Path.Combine(projectDirectory, containingFile);
 
                if (containingFile != binaryName)
                    System.IO.File.Copy(containingFile, binaryName);

                itemNode.Attributes["ContainingFile"].InnerText = "{" + itemNode.Attributes["Key"].InnerText + "}" + oldFileName;

                if (helpFile != helpName)
                {
                    if (true == System.IO.File.Exists(helpFile))
                    {
                        System.IO.File.Copy(helpFile, helpName);
                        itemNode.Attributes["HelpFile"].InnerText = "{" + itemNode.Attributes["Key"].InnerText + "}" + helpFileName;
                    }
                    else
                        itemNode.Attributes["HelpFile"].InnerText = "";
                }
            }

            if (true == System.IO.File.Exists(fileName))
                System.IO.File.Delete(fileName);

            _document.Save(fileName);
        }

        #endregion

        #region LoadMethods    

        private void LoadInterfaces(COMComponentReaderSettings settings, bool wantDispatch)
        {
            string xpathInterfaces = "";
            if (true == wantDispatch)
                xpathInterfaces = "DispatchInterfaces";
            else
                xpathInterfaces = "Interfaces";

            #region LoadInterfaces

            foreach (ComponentFile itemFile in settings.Files)
            {
                if (false == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                    string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });
                    XmlNode projectNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project" + formula);
                    formula = Utils.ChildNodeParamExpression(new string[] { "@GUID", XmlConvert.EncodeName(libInfo.GUID.Replace("{", "").Replace("}", "")), "@Name", libInfo.Name, "@MajorVersion", libInfo.MajorVersion.ToString(), "@MinorVersion", libInfo.MinorVersion.ToString() });
                    XmlNode componentNode = _document.SelectSingleNode(XPathConstants.Components + "/Component" + formula);

                    XmlNode interfacesNode = projectNode.SelectSingleNode(xpathInterfaces);
                    TLI.Interfaces interfaces = libInfo.Interfaces;
                    foreach (TLI.InterfaceInfo itemInterface in interfaces)
                    {
                        if (true == IsTargetInterfaceType(itemInterface.TypeKind, wantDispatch))
                        {
                            #region Create

                            formula = Utils.AttributeParamExpression(new string[] { "Name", itemInterface.Name });
                            XmlNode interfaceNode = interfacesNode.SelectSingleNode("Interface" + formula);
                            if (null == interfaceNode)
                            {
                                interfaceNode = _document.CreateElement("Interface");
                                Utils.AddAtrributeToNode(interfaceNode, "Name", itemInterface.Name);
                                Utils.AddAtrributeToNode(interfaceNode, "Caption", itemInterface.Name);
                                Utils.AddAtrributeToNode(interfaceNode, "Key", XmlConvert.EncodeName(Guid.NewGuid().ToString()));

                                interfaceNode.AppendChild(_document.CreateElement("Methods"));
                                interfaceNode.AppendChild(_document.CreateElement("Properties"));
                                interfaceNode.AppendChild(_document.CreateElement("RefComponents"));
                                interfaceNode.AppendChild(_document.CreateElement("GUIDs"));
                                interfaceNode.AppendChild(_document.CreateElement("InheritedInterfaces"));

                                interfacesNode.AppendChild(interfaceNode);
                            }

                            #endregion

                            #region RefComponent

                            XmlNode refComponentNode = interfaceNode.SelectSingleNode("RefComponents/Component[@Key='" + componentNode.Attributes["Key"].InnerText + "']");
                            if (null == refComponentNode)
                            {
                                refComponentNode = _document.CreateElement("RefComponent");
                                Utils.AddAtrributeToNode(refComponentNode, "Key", componentNode.Attributes["Key"].InnerText);
                                interfaceNode.SelectSingleNode("RefComponents").AppendChild(refComponentNode);
                            }
                            #endregion

                            #region Guid

                            XmlNode refGuid = interfaceNode.SelectSingleNode("GUIDs/Guid[@ComponentKey='" + componentNode.Attributes["Key"].InnerText + "']");
                            if (null == refGuid)
                            {
                                refGuid = _document.CreateElement("Guid");
                                Utils.AddAtrributeToNode(refGuid, "Guid", XmlConvert.EncodeName(itemInterface.GUID.Replace("{", "").Replace("}", "")));
                                Utils.AddAtrributeToNode(refGuid, "ComponentKey", componentNode.Attributes["Key"].InnerText);

                                interfaceNode.SelectSingleNode("GUIDs").AppendChild(refGuid);
                            }

                            #endregion

                            #region Methods & Properties

                            foreach (TLI.MemberInfo itemMember in itemInterface.Members)
                            {
                                if (true == IsInterfaceMethod(itemMember))
                                {
                                    _methodHandler.AddMethod(componentNode, projectNode, interfaceNode, itemMember);
                                }
                                else if (true == IsInterfaceProperty(itemMember))
                                {
                                    _propertyHandler.AddProperty(componentNode, projectNode, interfaceNode, itemMember);
                                }
                            }

                            #endregion
                        }

                        Marshal.ReleaseComObject(itemInterface);
                    }

                    Marshal.ReleaseComObject(interfaces);
                    Marshal.ReleaseComObject(libInfo);
                }
            }

            #endregion

            #region Scan Inherited Interfaces

            foreach (ComponentFile itemFile in settings.Files)
            {
                if (false == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                    string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });
                    XmlNode projectNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project" + formula);
                    XmlNode interfacesNode = projectNode.SelectSingleNode(xpathInterfaces);
                    TLI.Interfaces interfaces = libInfo.Interfaces;
                    foreach (TLI.InterfaceInfo itemInterface in interfaces)
                    {
                        if (true == IsTargetInterfaceType(itemInterface.TypeKind, wantDispatch))
                        {
                            formula = Utils.AttributeParamExpression(new string[] { "Name", itemInterface.Name });
                            XmlNode interfaceNode = interfacesNode.SelectSingleNode("Interface" + formula);

                            #region ImpliedInterfaces

                            if (null != itemInterface.VTableInterface)
                            {
                                TLI.Interfaces impliedInterfaces = itemInterface.VTableInterface.ImpliedInterfaces;
                                foreach (TLI.InterfaceInfo itemImplied in impliedInterfaces)
                                {
                                    if (("IDispatch" != itemImplied.Name) && ("IUnknown" != itemImplied.Name))
                                    {
                                        string key = "";

                                        XmlNode itemNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project/DispatchInterfaces/Interface[@Name='" + "" + itemImplied.Name + "']");
                                        if (null != itemNode)
                                        {
                                            key = itemNode.Attributes["Key"].InnerText;
                                            XmlNode refImplied = interfaceNode.SelectSingleNode("InheritedInterfaces/RefInterface[@Key='" + key + "']");
                                            if (null == refImplied)
                                            {
                                                refImplied = _document.CreateElement("RefInterface");
                                                Utils.AddAtrributeToNode(refImplied, "Key", key);
                                                interfaceNode.SelectSingleNode("InheritedInterfaces").AppendChild(refImplied);
                                            }
                                        }
                                        else
                                        {
                                            itemNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project/Interfaces/Interface[@Name='" + "" + itemImplied.Name + "']");
                                            if (null != itemNode)
                                            {
                                                key = itemNode.Attributes["Key"].InnerText;
                                                XmlNode refImplied = interfaceNode.SelectSingleNode("InheritedInterfaces/RefInterface[@Key='" + key + "']");
                                                if (null == refImplied)
                                                {
                                                    refImplied = _document.CreateElement("RefInterface");
                                                    Utils.AddAtrributeToNode(refImplied, "Key", key);
                                                    interfaceNode.SelectSingleNode("InheritedInterfaces").AppendChild(refImplied);
                                                }
                                            }
                                        }


                                    }
                                }

                                Marshal.ReleaseComObject(impliedInterfaces);

                            }

                            #endregion
                        }
                    }

                    Marshal.ReleaseComObject(interfaces);
                    Marshal.ReleaseComObject(libInfo);
                }
            }

            #endregion

            #region Remove Defined Methods and Properties from Interfaces there defined in Inherited
            {
               
                foreach (ComponentFile itemFile in settings.Files)
                {
                    if (false == itemFile.IsExternal)
                    { 
                        TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                        string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });
                        XmlNode projectNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project" + formula);
                        XmlNode interfacesNode = projectNode.SelectSingleNode(xpathInterfaces);
                        foreach(XmlNode interfaceNode in interfacesNode.ChildNodes)
                        {
                            XmlNodeList inheritedList = interfaceNode.SelectNodes("InheritedInterfaces/RefInterface");
                            foreach (XmlNode inheritedRefNode in inheritedList)
                            {
                                XmlNode inheritedDefinitionNode = GetInterfaceNode(inheritedRefNode);
                                RemoveInheritedMember(interfaceNode, inheritedDefinitionNode); 
                            }

                        }
                        Marshal.ReleaseComObject(libInfo);
                    }
                }

            }
            #endregion
        }

        
        List<XmlNode> GetInheritedInterfaces(XmlNode interfaceNode, List<XmlNode> listInherited)
        {
            XmlNodeList inheritedList = interfaceNode.SelectNodes("InheritedInterfaces/RefInterface");
            foreach (XmlNode inheritedRefNode in inheritedList)
            {
                XmlNode inheritedDefinitionNode = GetInterfaceNode(inheritedRefNode);
                GetInheritedInterfaces(inheritedDefinitionNode, listInherited);
            }

            return listInherited;
        }

        List<XmlNode> GetInheritedInterfaces(XmlNode interfaceNode)
        {
            List<XmlNode> listInherited = new List<XmlNode>();
            XmlNodeList inheritedList = interfaceNode.SelectNodes("InheritedInterfaces/RefInterface");
            foreach (XmlNode inheritedRefNode in inheritedList)
            {
                XmlNode inheritedDefinitionNode = GetInterfaceNode(inheritedRefNode);
                listInherited.Add(inheritedDefinitionNode);
                GetInheritedInterfaces(inheritedDefinitionNode, listInherited);
            }

            return listInherited;
        }

        private void RemoveInheritedMember(XmlNode interfaceNode, XmlNode inheritedInferfaceNode)
        {
            List<XmlNode> listToDelete = new List<XmlNode>();

            foreach (XmlNode methodNode in interfaceNode.SelectSingleNode("Methods"))
            {   
                string methodName = methodNode.Attributes["Name"].InnerText;
                XmlNode inheritedMethod = inheritedInferfaceNode.SelectSingleNode("Methods/Method[@Name ='" + methodName + "']");
                if (null != inheritedMethod)
                    listToDelete.Add(methodNode);
            }

            foreach (XmlNode propertyNode in interfaceNode.SelectSingleNode("Properties"))
            {
                string methodName = propertyNode.Attributes["Name"].InnerText;
                XmlNode inheritedMethod = inheritedInferfaceNode.SelectSingleNode("Properties/Property[@Name ='" + methodName + "']");
                if (null != inheritedMethod)
                    listToDelete.Add(propertyNode);
            }

            foreach (XmlNode nodeToDelete in listToDelete)
            {
                XmlNode parentNode = nodeToDelete.ParentNode;
                parentNode.RemoveChild(nodeToDelete);
            }
         
            List<XmlNode> inheritedList = GetInheritedInterfaces(inheritedInferfaceNode);
           
            foreach (XmlNode inheritedRefNode in inheritedList)
            {
                XmlNode inheritedDefinitionNode = GetInterfaceNode(inheritedRefNode);
                GetInheritedInterfaces(inheritedDefinitionNode, inheritedList);
                RemoveInheritedMember(interfaceNode, inheritedDefinitionNode);
            }


           // RemoveInheritedMember(inheritedInferfaceNode, inheritedList);
        }


        private XmlNode GetInterfaceNode(XmlNode refInterface)
        {
            string key = refInterface.Attributes["Key"].InnerText;
            XmlNode node = refInterface.OwnerDocument.SelectSingleNode(XPathConstants.Solution + "/Projects/Project/DispatchInterfaces/Interface[@Key ='" + key + "']");
            if (node != null)
                return node;

            node = refInterface.OwnerDocument.SelectSingleNode(XPathConstants.Solution + "/Projects/Project/Interfaces/Interface[@Key ='" + key + "']");
            if (node != null)
                return node;

            throw (new ArgumentException("Interface not found" + key));
        }

        private void LoadCoClasses(COMComponentReaderSettings settings)
        {
            foreach (ComponentFile itemFile in settings.Files)
            {
                if (false == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                    string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });
                    XmlNode projectNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project" + formula);
                    formula = Utils.ChildNodeParamExpression(new string[] { "@GUID", XmlConvert.EncodeName(libInfo.GUID.Replace("{", "").Replace("}", "")), "@Name", libInfo.Name, "@MajorVersion", libInfo.MajorVersion.ToString(), "@MinorVersion", libInfo.MinorVersion.ToString() });
                    XmlNode componentNode = _document.SelectSingleNode(XPathConstants.Components + "/Component" + formula);

                    XmlNode classesNode = projectNode.SelectSingleNode("CoClasses");
                    TLI.CoClasses classes = libInfo.CoClasses;
                    foreach (TLI.CoClassInfo itemClass in classes)
                    {
                        #region Create 

                        formula = Utils.AttributeParamExpression(new string[] { "Name", itemClass.Name });
                        XmlNode classNode = classesNode.SelectSingleNode("CoClass" + formula);
                        
                        if (null == classNode)
                        {
                            classNode = _document.CreateElement("CoClass");
                            Utils.AddAtrributeToNode(classNode, "Name", itemClass.Name);
                            Utils.AddAtrributeToNode(classNode, "Caption", itemClass.Name);
                            classNode.AppendChild(_document.CreateElement("RefComponents"));
                            classNode.AppendChild(_document.CreateElement("GUIDs"));
                            classNode.AppendChild(_document.CreateElement("Properties"));
                            classNode.AppendChild(_document.CreateElement("Methods"));
                            classNode.AppendChild(_document.CreateElement("InheritedInterfaces"));
                            classesNode.AppendChild(classNode);
                        }

                        #endregion

                        #region RefComponent

                        XmlNode refComponentNode = classNode.SelectSingleNode("RefComponents/Component[@Key='" + componentNode.Attributes["Key"].InnerText + "']");
                        if (null == refComponentNode)
                        {
                            refComponentNode = _document.CreateElement("RefComponent");
                            Utils.AddAtrributeToNode(refComponentNode, "Key", componentNode.Attributes["Key"].InnerText);
                            classNode.SelectSingleNode("RefComponents").AppendChild(refComponentNode);
                        }

                        #endregion

                        #region Guid

                        XmlNode refGuid = classNode.SelectSingleNode("GUIDs/Guid[@ComponentKey='" + componentNode.Attributes["Key"].InnerText + "']");
                        if (null == refGuid)
                        {
                            refGuid = _document.CreateElement("Guid");
                            Utils.AddAtrributeToNode(refGuid, "Guid", XmlConvert.EncodeName(itemClass.GUID.Replace("{", "").Replace("}", "")));
                            Utils.AddAtrributeToNode(refGuid, "ComponentKey", componentNode.Attributes["Key"].InnerText);

                            classNode.SelectSingleNode("GUIDs").AppendChild(refGuid);
                        }

                        #endregion

                        #region Methods & Properties

                        foreach (TLI.MemberInfo itemMember in itemClass.DefaultInterface.Members)
                        {
                            if (true == IsInterfaceMethod(itemMember))
                            {
                                _methodHandler.AddMethod(componentNode, projectNode, classNode, itemMember);
                            }
                            else if (true == IsInterfaceProperty(itemMember))
                            {
                                _propertyHandler.AddProperty(componentNode, projectNode, classNode, itemMember);
                            }
                        }

                        #endregion

                        #region EventInterface

                        InterfaceInfo eventInterface = itemClass.DefaultEventInterface;
                        if (null != eventInterface)
                        {
                            string key = GetEventInterfaceKey(classNode, eventInterface);
                            XmlNode eventInterfaceNode = classNode.AppendChild(_document.CreateElement("EventInterface"));
                            Utils.AddAtrributeToNode(eventInterfaceNode, "Key", key);
                            Marshal.ReleaseComObject(eventInterface);
                        }

                        #endregion

                        Marshal.ReleaseComObject(itemClass);
                    }
                    
                    Marshal.ReleaseComObject(classes);
                    Marshal.ReleaseComObject(libInfo);
                }
            }

            foreach (ComponentFile itemFile in settings.Files)
            {
                if (false == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                    string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });
                    XmlNode projectNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project" + formula);
                    XmlNode classesNode = projectNode.SelectSingleNode("CoClasses");
                    TLI.CoClasses coClasses = libInfo.CoClasses;
                    foreach (TLI.CoClassInfo itemClass in coClasses)
                    {
                            formula = Utils.AttributeParamExpression(new string[] { "Name", itemClass.Name });
                            XmlNode interfaceNode = classesNode.SelectSingleNode("CoClass" + formula);

                            #region ImpliedInterfaces
                          
                            TLI.Interfaces impliedInterfaces = itemClass.Interfaces;
                            if (null != impliedInterfaces)
                            {
                                foreach (TLI.InterfaceInfo itemImplied in itemClass.Interfaces)
                                {
                                    string eventInterfaceName = "";
                                    if (null != itemClass.DefaultEventInterface)
                                        eventInterfaceName = itemClass.DefaultEventInterface.Name;

                                    if (("IDispatch" != itemImplied.Name) && ("IUnknown" != itemImplied.Name) && (itemImplied.Name != eventInterfaceName))
                                    {
                                        string key = "";

                                        XmlNode itemNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project/DispatchInterfaces/Interface[@Name='" + "" + itemImplied.Name + "']");
                                        if (null != itemNode)
                                        {
                                            key = itemNode.Attributes["Key"].InnerText;
                                            XmlNode refImplied = interfaceNode.SelectSingleNode("InheritedInterfaces/RefInterface[@Key='" + key + "']");
                                            if (null == refImplied)
                                            {
                                                refImplied = _document.CreateElement("RefInterface");
                                                Utils.AddAtrributeToNode(refImplied, "Key", key);
                                                interfaceNode.SelectSingleNode("InheritedInterfaces").AppendChild(refImplied);
                                            }
                                        }
                                        else
                                        {
                                            itemNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project/Interfaces/Interface[@Name='" + "" + itemImplied.Name + "']");
                                            if (null != itemNode)
                                            {
                                                key = itemNode.Attributes["Key"].InnerText;
                                                XmlNode refImplied = interfaceNode.SelectSingleNode("InheritedInterfaces/RefInterface[@Key='" + key + "']");
                                                if (null == refImplied)
                                                {
                                                    refImplied = _document.CreateElement("RefInterface");
                                                    Utils.AddAtrributeToNode(refImplied, "Key", key);
                                                    interfaceNode.SelectSingleNode("InheritedInterfaces").AppendChild(refImplied);
                                                }
                                            }
                                        }


                                    }
                                }

                                Marshal.ReleaseComObject(impliedInterfaces);

                            }

                            #endregion
                    }

                    Marshal.ReleaseComObject(coClasses);
                    Marshal.ReleaseComObject(libInfo);
                }
            }
 
        }

        private void LoadEnums(COMComponentReaderSettings settings)
        {
            foreach (ComponentFile itemFile in settings.Files)
            {
                if (false == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                    string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });
                    XmlNode projectNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project" + formula);
                    
                    formula = Utils.ChildNodeParamExpression(new string[] { "@GUID", XmlConvert.EncodeName(libInfo.GUID.Replace("{", "").Replace("}", "")), "@Name", libInfo.Name, "@MajorVersion", libInfo.MajorVersion.ToString(), "@MinorVersion", libInfo.MinorVersion.ToString() });
                    XmlNode componentNode = _document.SelectSingleNode(XPathConstants.Components + "/Component" + formula);
                    
                    XmlNode enumsNode = projectNode.SelectSingleNode("Enums");
                    TLI.Constants constants = libInfo.Constants;
                    foreach (TLI.ConstantInfo itemConstant in constants)
                    {
                        formula = Utils.AttributeParamExpression(new string[] { "Name", itemConstant.Name });
                        XmlNode enumNode = enumsNode.SelectSingleNode("Enum" + formula);
                        if (null == enumNode)
                        {
                            enumNode = _document.CreateElement("Enum");
                            Utils.AddAtrributeToNode(enumNode, "Name", itemConstant.Name);
                            enumNode.AppendChild(_document.CreateElement("RefComponents"));
                            enumNode.AppendChild(_document.CreateElement("EnumMembers"));
                            enumsNode.AppendChild(enumNode);
                        }
                       
                        XmlNode refComponentNode = enumNode.SelectSingleNode("RefComponents/Component[@Key='" + componentNode.Attributes["Key"].InnerText + "']");
                        if (null == refComponentNode)
                        {
                            refComponentNode = _document.CreateElement("RefComponent");
                            Utils.AddAtrributeToNode(refComponentNode, "Key", componentNode.Attributes["Key"].InnerText);
                            enumNode.SelectSingleNode("RefComponents").AppendChild(refComponentNode);
                        }

                        TLI.Members members = itemConstant.Members;
                        foreach (TLI.MemberInfo itemMember in members)
                        {
                            XmlNode enumMemberNode = enumNode.SelectSingleNode("EnumMembers/EnumMember[@Name='" + itemMember.Name + "']");
                            if (null == enumMemberNode)
                            {
                                enumMemberNode = _document.CreateElement("EnumMember");
                                Utils.AddChildToNode(enumMemberNode, "RefComponents");
                                Utils.AddAtrributeToNode(enumMemberNode, "Name", itemMember.Name);
                                Utils.AddAtrributeToNode(enumMemberNode, "Value", itemMember.Value.ToString());
                                enumNode.SelectSingleNode("EnumMembers").AppendChild(enumMemberNode);
                            }

                            refComponentNode = enumMemberNode.SelectSingleNode("RefComponents/Component[@Key='" + componentNode.Attributes["Key"].InnerText + "']");
                            if (null == refComponentNode)
                            {
                                refComponentNode = _document.CreateElement("RefComponent");
                                Utils.AddAtrributeToNode(refComponentNode, "Key", componentNode.Attributes["Key"].InnerText);
                                enumMemberNode.SelectSingleNode("RefComponents").AppendChild(refComponentNode);
                            }

                            Marshal.ReleaseComObject(itemMember);
                        }
                        Marshal.ReleaseComObject(members);

                        
                        Marshal.ReleaseComObject(itemConstant);
                    }
                    Marshal.ReleaseComObject(constants);

                    Marshal.ReleaseComObject(libInfo);
                }
            }
        }

        private void LoadSolution(COMComponentReaderSettings settings)
        {
            foreach (ComponentFile itemFile in settings.Files)
            {
                if (false == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);

                    string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });
                    XmlNode projectNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project" + formula);
                    if (null == projectNode)
                    {   
                        projectNode = _document.CreateElement("Project");
                        Utils.AddAtrributeToNode(projectNode, "Name", libInfo.Name);
                        Utils.AddAtrributeToNode(projectNode, "ProjectName", "LateBindingApi." + libInfo.Name);
                        Utils.AddAtrributeToNode(projectNode, "Namespace", "LateBindingApi." + libInfo.Name);
                        Utils.AddAtrributeToNode(projectNode, "Prefix", libInfo.Name.Substring(0, 2).ToUpper());
                        Utils.AddAtrributeToNode(projectNode, "Key", XmlConvert.EncodeName(Guid.NewGuid().ToString()));

                        Utils.AddChildToNode(projectNode, "Enums");
                        Utils.AddChildToNode(projectNode, "CoClasses");
                        Utils.AddChildToNode(projectNode, "DispatchInterfaces");
                        Utils.AddChildToNode(projectNode, "Interfaces");
                        Utils.AddChildToNode(projectNode, "RefComponents");
                        
                        _document.SelectSingleNode(XPathConstants.Solution + "/Projects").AppendChild(projectNode);
                    }

                    formula = Utils.ChildNodeParamExpression(new string[] { "@GUID", XmlConvert.EncodeName(libInfo.GUID.Replace("{", "").Replace("}", "")), "@Name", libInfo.Name, "@MajorVersion", libInfo.MajorVersion.ToString(), "@MinorVersion", libInfo.MinorVersion.ToString() });
                    XmlNode componentNode = _document.SelectSingleNode(XPathConstants.Components + "/Component" + formula);
                    XmlNode refComponentsNode = projectNode.SelectSingleNode("RefComponents");
                    XmlNode refComponentNode = Utils.AddChildToNode(refComponentsNode, "RefComponent");
                    Utils.AddAtrributeToNode(refComponentNode, "Key", componentNode.Attributes["Key"].InnerText);

                    Marshal.ReleaseComObject(libInfo);
                }
            }
        }

        private void LoadExternalsComponents(COMComponentReaderSettings settings)
        {          
            foreach (ComponentFile itemFile in settings.Files)
            {
                if (true == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                    string formula = Utils.AttributeParamExpression(new string[] { "Name", libInfo.Name });

                    XmlNode externalNode = _document.SelectSingleNode(XPathConstants.Externals + "/External" + formula);
                    if (null == externalNode)
                    {
                       externalNode = _document.CreateElement("External");
                       Utils.AddChildToNode(externalNode, "ExternalComponents");
                       Utils.AddAtrributeToNode(externalNode, "Name", libInfo.Name);
                       Utils.AddAtrributeToNode(externalNode, "Prefix", libInfo.Name.Substring(0,2).ToUpper());
                       Utils.AddAtrributeToNode(externalNode, "Namespace", "LateBindingApi." + libInfo.Name);
                       Utils.AddAtrributeToNode(externalNode, "Key", XmlConvert.EncodeName(Guid.NewGuid().ToString()));
                       _document.SelectSingleNode(XPathConstants.Externals).AppendChild(externalNode);
                    }
                    XmlNode externalComponentsNode = externalNode.SelectSingleNode("ExternalComponents");

                    formula = Utils.ChildNodeParamExpression(new string[] { "@GUID", XmlConvert.EncodeName(libInfo.GUID), "@MajorVersion", libInfo.MajorVersion.ToString(), "@MinorVersion", libInfo.MinorVersion.ToString() });
                    XmlNode externalComponentNode = externalComponentsNode.SelectSingleNode("ExternalComponent" + formula);
                    if (null == externalComponentNode)
                    {
                        externalComponentNode = _document.CreateElement("ExternalComponent");
                        Utils.AddAtrributeToNode(externalComponentNode, "Key", XmlConvert.EncodeName(Guid.NewGuid().ToString()));
                        Utils.AddAtrributeToNode(externalComponentNode, "GUID", XmlConvert.EncodeName(libInfo.GUID.Replace("{", "").Replace("}", "")));
                        Utils.AddAtrributeToNode(externalComponentNode, "MajorVersion", libInfo.MajorVersion.ToString());
                        Utils.AddAtrributeToNode(externalComponentNode, "MinorVersion", libInfo.MinorVersion.ToString());
                        Utils.AddAtrributeToNode(externalComponentNode, "Description", Utils.GetTypeLibDescription(libInfo.MajorVersion.ToString(), libInfo.MinorVersion.ToString(), libInfo.GUID));
                        Utils.AddAtrributeToNode(externalComponentNode, "SysKind", Convert.ToInt32(libInfo.SysKind).ToString());
                        externalComponentsNode.AppendChild(externalComponentNode);
                    }
                    Marshal.ReleaseComObject(libInfo);
                }
            }
        }

        private void LoadComponents(COMComponentReaderSettings settings)
        {
            int i = 1;
            foreach (ComponentFile itemFile in settings.Files)
            {
                if (false == itemFile.IsExternal)
                {
                    TypeLibInfo libInfo = _typeLibApplication.TypeLibInfoFromFile(itemFile.Filename);
                    string formula = Utils.ChildNodeParamExpression(new string[] { "@GUID", XmlConvert.EncodeName(libInfo.GUID), "@Name", libInfo.Name, "@MajorVersion", libInfo.MajorVersion.ToString(), "@MinorVersion", libInfo.MinorVersion.ToString() });
                    XmlNode componentNode = _document.SelectSingleNode(XPathConstants.Components + "/Component" + formula);
                    if (null == componentNode)
                    {
                        componentNode = _document.CreateElement("Component");

                        Utils.AddAtrributeToNode(componentNode, "Name", libInfo.Name);
                        Utils.AddAtrributeToNode(componentNode, "ContainingFile", libInfo.ContainingFile);
                        Utils.AddAtrributeToNode(componentNode, "Key", XmlConvert.EncodeName(Guid.NewGuid().ToString()));
                        Utils.AddAtrributeToNode(componentNode, "GUID", XmlConvert.EncodeName(libInfo.GUID.Replace("{", "").Replace("}", "")));
                        Utils.AddAtrributeToNode(componentNode, "HelpFile",libInfo.HelpFile);
                        Utils.AddAtrributeToNode(componentNode, "MajorVersion", libInfo.MajorVersion.ToString());
                        Utils.AddAtrributeToNode(componentNode, "MinorVersion", libInfo.MinorVersion.ToString());
                        Utils.AddAtrributeToNode(componentNode, "LCID", libInfo.LCID.ToString());
                        Utils.AddAtrributeToNode(componentNode, "Description", Utils.GetTypeLibDescription(libInfo.GUID, libInfo.ToString(), libInfo.MinorVersion.ToString()));
                        Utils.AddAtrributeToNode(componentNode, "Version", libInfo.Name.Substring(0, 2).ToUpper() + i.ToString());
                        Utils.AddAtrributeToNode(componentNode, "SysKind", Convert.ToInt32(libInfo.SysKind).ToString());

                        _document.SelectSingleNode(XPathConstants.Components).AppendChild(componentNode);
                    }
                    i++;
                }
            }
        }

        #endregion

        #region ResetDocument

        private void ResetDocument()
        {
            if (null == _document)
                _document = new XmlDocument();

            _document.RemoveAll();
            
            string xml = Utils.ReadTextFileFromRessource("LateBindingApi.CodeGenerator.Document.xml");
            _document.LoadXml(xml);

            if (_document.Schemas.Count == 0)
            {
                Stream xsdStream = Utils.ReadStreamFromRessource("LateBindingApi.CodeGenerator.Document.xsd");
                Utils.AddSchema(_document, xsdStream);
                xsdStream.Close();
                List<ValidationEventArgs> errorList = Utils.Validate(_document);
                if (errorList.Count > 0)
                {
                    string message = string.Format("XML Template contains {0} schema errors.", errorList.Count);
                    throw (new XmlException(message));
                }
            }
        }

        #endregion 

        #region VarType Detection
 
        internal bool IsComProxy(VarTypeInfo typeInfo)
        {
            if (TliVarType.VT_DISPATCH == typeInfo.VarType)
                return true;

            if(null == typeInfo.TypeInfo)
                return false;

            if ((typeInfo.TypeInfo.TypeKind == TypeKinds.TKIND_DISPATCH) || (typeInfo.TypeInfo.TypeKind == TypeKinds.TKIND_INTERFACE)
            || (typeInfo.TypeInfo.TypeKind == TypeKinds.TKIND_COCLASS))
            {
                return true;
            }
            else
            {
                 return false;
            }
        }

        internal string GetEventInterfaceKey(XmlNode classNode, InterfaceInfo faceInfo)
        {
           
            XmlNode faceNode = classNode.ParentNode.ParentNode.SelectSingleNode("Interfaces/Interface[@Name='" + faceInfo.Name + "']");
            if (null != faceNode)
            {
                return faceNode.Attributes["Key"].InnerText;
            }

            faceNode = classNode.ParentNode.ParentNode.SelectSingleNode("DispatchInterfaces/Interface[@Name='" + faceInfo.Name + "']");
            if (null != faceNode)
            {
                return faceNode.Attributes["Key"].InnerText;
            }

            throw (new ArgumentException("Interface not found " + faceNode.Name));
        }

        internal string GetSourceKey(VarTypeInfo typeInfo)
        {            
           
            if (true == typeInfo.IsExternalType)
            {
                string targetGuid = XmlConvert.EncodeName(typeInfo.TypeLibInfoExternal.GUID.Replace("{", "").Replace("}", ""));
                XmlNode externalsNode = _document.SelectSingleNode(XPathConstants.Externals);
                XmlNode externalNode = externalsNode.SelectSingleNode("External/ExternalComponents/ExternalComponent[@GUID='" + targetGuid + "']");
                if (null != externalNode)
                {
                    return externalNode.ParentNode.ParentNode.Attributes["Key"].InnerText;
                }
                else
                    return "";
            }
            else
            {
                if (null == typeInfo.TypeInfo)
                    return "";

                string targetGuid = XmlConvert.EncodeName(typeInfo.TypeInfo.Parent.GUID.Replace("{", "").Replace("}", ""));
                XmlNode componentNode = _document.SelectSingleNode(XPathConstants.Components + "/Component[@GUID='" + targetGuid + "']");
                if (null != componentNode)
                {
                    string key = componentNode.Attributes["Key"].InnerText;
                    XmlNode refKeyNode = _document.SelectSingleNode(XPathConstants.Solution + "/Projects/Project/RefComponents/RefComponent[@Key='" + key + "']");
                    if (null != componentNode)
                    {
                        return refKeyNode.ParentNode.ParentNode.Attributes["Key"].InnerText;
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }
        }

        internal string ValidateType(VarTypeInfo typeInfo)
        {
            switch (typeInfo.VarType)
            {

                case TliVarType.VT_EMPTY:       // Type in Component
                    return typeInfo.TypeInfo.Name;
                case TliVarType.VT_I4:          // Int32
                    return typeInfo.TypedVariant.GetType().Name;
                case TliVarType.VT_INT:         // Int32
                    return typeInfo.TypedVariant.GetType().Name;
                case TliVarType.VT_R8:          // Double
                    return typeInfo.TypedVariant.GetType().Name;
                case TliVarType.VT_BSTR:        // String
                    return "string";
                case TliVarType.VT_BOOL:        // Bool
                    return "bool";
                case TliVarType.VT_DISPATCH:     // COMProxy
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_UNKNOWN:      // COMProxy
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_VOID: 
                    return "void";
                /*Unkown*/
                case TliVarType.VT_VARIANT:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_FILETIME:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_BLOB:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_STREAM:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_STORAGE:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_STREAMED_OBJECT:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_STORED_OBJECT:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_BLOB_OBJECT:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_CF:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_CLSID:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_VECTOR:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_ARRAY:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_BYREF:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_RESERVED:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_HRESULT:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_PTR:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_SAFEARRAY:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_CARRAY:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_USERDEFINED:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_LPSTR:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_LPWSTR:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_RECORD:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_I1:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_I8:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_UI1:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_UI2:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_UI4:
                    return typeInfo.TypedVariant.GetType().Name;
                case TliVarType.VT_UI8:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_UINT:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_DECIMAL:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_ERROR:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_CY:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_DATE:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_R4:
                    return typeInfo.VarType.ToString();
                case TliVarType.VT_I2:
                    return typeInfo.VarType.ToString();
                default:
                    return typeInfo.VarType.ToString();
                //throw (new ArgumentException("Unkown ReturnType " + typeInfo.VarType.ToString()));
            }
        }

        internal string GetReturnType(MemberInfo methodInfo)
        {
            return ValidateType(methodInfo.ReturnType);
        }

        internal string GetPropertyType(MemberInfo propertyInfo)
        {
            return ValidateType(propertyInfo.ReturnType);
        }

        internal string GetParameterType(ParameterInfo paramInfo)
        {
            return ValidateType(paramInfo.VarTypeInfo);
        }

        internal bool IsIDispatchOrIUnkownMethod(TLI.MemberInfo memberInfo)
        {
            switch (memberInfo.Name)
            {
                case "GetTypeInfoCount":
                case "GetTypeInfo":
                case "GetIDsOfNames":
                case "Invoke":
                    return true;
            }

            switch (memberInfo.Name)
            {
                case "AddRef":
                case "Release":
                case "QueryInterface":
                    return true;
            }

            return false;
        }

        #endregion

        #region Private Helper

        private bool IsTargetInterfaceType(TypeKinds kind, bool wantDispatch)
        {
            if (true == wantDispatch)
            {
                if (kind == TypeKinds.TKIND_DISPATCH)
                    return true;
                else
                    return false;

            }
            else
            {
                if (kind == TypeKinds.TKIND_DISPATCH)
                    return false;
                else
                    return true;
            }
        }

        private bool IsInterfaceMethod(TLI.MemberInfo memberInfo)
        {
            bool isImpliedComMethod = IsIDispatchOrIUnkownMethod(memberInfo);
            if (true == isImpliedComMethod)
                return false;

            if (memberInfo.DescKind == DescKinds.DESCKIND_FUNCDESC)
            {
                if ((memberInfo.InvokeKind != InvokeKinds.INVOKE_PROPERTYGET) &&
                    (memberInfo.InvokeKind != InvokeKinds.INVOKE_PROPERTYPUT) &&
                    (memberInfo.InvokeKind != InvokeKinds.INVOKE_PROPERTYPUTREF))
                    return true;
                else
                    return false;
            }
            else
            {
                if (memberInfo.InvokeKind == InvokeKinds.INVOKE_FUNC)
                    return true;
                else
                    return false;
            }

        }

        private bool IsInterfaceProperty(TLI.MemberInfo memberInfo)
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

        #endregion
    }
}
