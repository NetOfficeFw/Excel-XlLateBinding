using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.Visual
{
    public partial class ComponentTreeView : UserControl
    {
        private string _currentNodePath;

        public event TreeViewEventHandler AfterSelect;

        public ComponentTreeView()
        {
            InitializeComponent();
        }

        public void Clear()
        {
            treeViewComponents.Nodes.Clear();
        }

        public void Show(XmlNode documentNode)
        {
            Clear();            
            treeViewComponents.Tag = documentNode;
  
            #region Components

            TreeNode tnComponents = treeViewComponents.Nodes.Add("Components");
            tnComponents.Tag = documentNode.SelectSingleNode("Components");
            XmlNodeList components = documentNode.SelectNodes("Components/Component");
            foreach (XmlNode componentNode in components)
            {
                string name = componentNode.Attributes["Name"].InnerText;
                TreeNode tnComponent = tnComponents.Nodes.Add(name);
                tnComponent.Tag = componentNode;
            }

            #endregion

            #region Externals

            TreeNode tnExternals = treeViewComponents.Nodes.Add("Externals");
            tnExternals.Tag = documentNode.SelectSingleNode("Externals");
            XmlNodeList externals = documentNode.SelectNodes("Externals/External");
            foreach (XmlNode externalNode in externals)
            {
                string name = externalNode.Attributes["Name"].InnerText;
                TreeNode tnExternal = tnExternals.Nodes.Add(name);
                tnExternal.Tag = externalNode;

                TreeNode tnExternalComponents = tnExternal.Nodes.Add("Components");
                tnExternalComponents.Tag = externalNode.SelectSingleNode("ExternalComponents");
                foreach (XmlNode exComponent in externalNode.SelectSingleNode("ExternalComponents"))
                {
                    string exName = exComponent.Attributes["MajorVersion"].InnerText + "." + exComponent.Attributes["MinorVersion"].InnerText;
                    TreeNode tnExComponent = tnExternalComponents.Nodes.Add(exName);
                    tnExComponent.Tag = exComponent;
                }

            }

            #endregion

            #region Projects

            TreeNode tnSolution = treeViewComponents.Nodes.Add("Solution");
            tnSolution.Tag = documentNode.SelectSingleNode("Solution");
            XmlNodeList projects = documentNode.SelectNodes("Solution/Projects/Project");
            foreach (XmlNode projectNode in projects)
            {
                string name = projectNode.Attributes["Name"].InnerText;
                TreeNode tnProject = tnSolution.Nodes.Add(name);
                tnProject.Tag = projectNode;

                #region Enums

                TreeNode tnEnums = tnProject.Nodes.Add("Enums");
                XmlNodeList enums = projectNode.SelectNodes("Enums/Enum");
                tnEnums.Tag = projectNode.SelectSingleNode("Enums");
                foreach (XmlNode enumNode in enums)
                {
                    string enumName = enumNode.Attributes["Name"].InnerText;

                    if (true == FilterPassed(enumName))
                    {
                        TreeNode tnEnum = tnEnums.Nodes.Add(enumName);
                        tnEnum.Tag = enumNode;
                    }
                    
                }

                #endregion

                #region CoClasses

                TreeNode tnClasses = tnProject.Nodes.Add("CoClasses");
                tnClasses.Tag = projectNode.SelectSingleNode("CoClasses");
                XmlNodeList classes = projectNode.SelectNodes("CoClasses/CoClass");
                foreach (XmlNode classNode in classes)
                {
                    string className = classNode.Attributes["Name"].InnerText;
                    if (true == FilterPassed(className))
                    {
                        TreeNode tnClass = tnClasses.Nodes.Add(className);
                        tnClass.Tag = classNode;

                        #region Inherited Interfaces

                        XmlNodeList refInterfaces = classNode.SelectNodes("InheritedInterfaces/RefInterface");
                        foreach (XmlNode itemRefInterface in refInterfaces)
                        {
                            XmlNode faceNode = GetInterfaceNode(itemRefInterface);
                            TreeNode tnRefInterface = tnClass.Nodes.Add(faceNode.Attributes["Name"].InnerText);
                            tnRefInterface.Tag = faceNode;
                        }

                        #endregion
                    }
                }

                #endregion

                #region Interfaces

                TreeNode tnInterfaces = tnProject.Nodes.Add("Interfaces");
                tnInterfaces.Tag = projectNode.SelectSingleNode("Interfaces");
                XmlNodeList interfaces = projectNode.SelectNodes("Interfaces/Interface");
                foreach (XmlNode interfaceNode in interfaces)
                {
                    string interfaceName = interfaceNode.Attributes["Name"].InnerText;
                    if (true == FilterPassed(interfaceName))
                    {
                        TreeNode tnInterface = tnInterfaces.Nodes.Add(interfaceName);
                        tnInterface.Tag = interfaceNode;

                        #region Inherited Interfaces

                        XmlNodeList refInterfaces = interfaceNode.SelectNodes("InheritedInterfaces/RefInterface");
                        foreach (XmlNode itemRefInterface in refInterfaces)
                        {
                            XmlNode faceNode = GetInterfaceNode(itemRefInterface);
                            TreeNode tnRefInterface = tnInterface.Nodes.Add(faceNode.Attributes["Name"].InnerText);
                            tnRefInterface.Tag = faceNode;
                        }

                        #endregion
 
                    }
                }

                #endregion

                #region DispatchInterfaces

                TreeNode tnDispatchInterfaces = tnProject.Nodes.Add("DispatchInterfaces");
                tnDispatchInterfaces.Tag = projectNode.SelectSingleNode("DispatchInterfaces");
                XmlNodeList dispatchInterfaces = projectNode.SelectNodes("DispatchInterfaces/Interface");
                foreach (XmlNode dispatchNode in dispatchInterfaces)
                {
                    string interfaceName = dispatchNode.Attributes["Name"].InnerText;
                    if (true == FilterPassed(interfaceName))
                    {
                        TreeNode tnInterface = tnDispatchInterfaces.Nodes.Add(interfaceName);
                        tnInterface.Tag = dispatchNode;

                        #region Inherited Interfaces

                        XmlNodeList refInterfaces = dispatchNode.SelectNodes("InheritedInterfaces/RefInterface");
                        foreach (XmlNode itemRefInterface in refInterfaces)
                        {
                            XmlNode faceNode = GetInterfaceNode(itemRefInterface);
                            TreeNode tnRefInterface = tnInterface.Nodes.Add(faceNode.Attributes["Name"].InnerText);
                            tnRefInterface.Tag = faceNode;
                        }

                        #endregion
                    }
                }

                #endregion
            }

            #endregion
        } 

        public void SelectFirstNode()
        {
            if(treeViewComponents.Nodes.Count > 0)
                treeViewComponents.SelectedNode = treeViewComponents.Nodes[0];
        }

        public void SaveCurrentNodePath()
        {
            if (null == treeViewComponents.SelectedNode)
                _currentNodePath = "";
            else
                _currentNodePath = treeViewComponents.SelectedNode.FullPath; 
        }

        public void RestoreExpandState()
        {
            TreeNode node = null;
            string[] splitArray = _currentNodePath.Split(new string[] { treeViewComponents.PathSeparator }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string nodeName in splitArray)
            {
                if(node == null)
                    node = SearchChildTree(treeViewComponents, nodeName);
                else
                    node = SearchChildTree(node,nodeName);
                
                if (node != null)
                node.Expand();                
            }
            treeViewComponents.SelectedNode =node;
        }

        private XmlNode GetInterfaceNode(XmlNode refInterfaceNode)
        {
            string key = refInterfaceNode.Attributes["Key"].InnerText;
            XmlNode interfaceNode = refInterfaceNode.OwnerDocument.DocumentElement.SelectSingleNode("Solution/Projects/Project/DispatchInterfaces/Interface[@Key ='" + key + "']");
            if (null != interfaceNode)
                return interfaceNode;

            interfaceNode = refInterfaceNode.OwnerDocument.DocumentElement.SelectSingleNode("Solution/Projects/Project/Interfaces/Interface[@Key ='" + key + "']");
            if (null != interfaceNode)
                return interfaceNode;

            throw (new ArgumentException("Interface not found " + key));
        }

        private TreeNode SearchChildTree(TreeView treeView, string name)
        {
            foreach (TreeNode tn in treeViewComponents.Nodes)
            {
                if (tn.Text == name)
                    return tn;
            }
            return null;
        } 

        private TreeNode SearchChildTree(TreeNode treeNode, string name)
        {            
            foreach (TreeNode tn in treeNode.Nodes)
            {
                if (tn.Text == name)
                    return tn;
            }
            return null;
        } 

        private bool FilterPassed(string expression)
        {
            string filterText = textBoxFilter.Text.Trim();
            if (filterText == "") return true;

            if (expression.IndexOf(filterText, 0, StringComparison.InvariantCultureIgnoreCase) > -1)
                return true;
            else
                return false;
        }

        private void textBoxFilter_KeyDown(object sender, KeyEventArgs e)
        {
            XmlNode documentNode = treeViewComponents.Tag as XmlNode;
            if ((null != documentNode) && (e.KeyCode == Keys.Return))
            {
                SaveCurrentNodePath();
                Show(documentNode);
                RestoreExpandState();
            }               
        }

        private void treeViewComponents_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (null != AfterSelect)
                AfterSelect(sender, e);
        }
    }
}
