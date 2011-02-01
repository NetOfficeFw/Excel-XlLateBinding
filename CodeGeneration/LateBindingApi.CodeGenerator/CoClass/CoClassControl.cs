using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using LateBindingApi.CodeGenerator.Core;

namespace LateBindingApi.CodeGenerator
{
    public partial class CoClassControl : UserControl
    {
        #region Public Methods
        
        public CoClassControl()
        {
            InitializeComponent();
            this.Visible = false;
            methodsControlMain.Initialize();
            propertiesControlMain.Initialize();
            eventsControlMain.Initialize();
            SetupCoClassComponentsGrid();
            SetupCoClassInterfacesGrid();
        }

        #endregion

        #region Public Methods

        public void Clear()
        {
            dataGridViewCoClassInterfaces.Rows.Clear();
            dataGridViewCoClassComponents.Rows.Clear();
            methodsControlMain.Clear();
            propertiesControlMain.Clear();
            eventsControlMain.Clear();
        }

        XmlNode GetEventInterfaceNode(XmlNode refNode)
        {
            string eventInterfaceName = refNode.Attributes["Name"].InnerText;
            XmlNode facesNode = refNode.OwnerDocument.SelectSingleNode(Constants.Interfaces);
            foreach (XmlNode interfaceNode in facesNode.ChildNodes )
            {
                if (interfaceNode.Attributes["Name"].InnerText == eventInterfaceName)
                    return interfaceNode;
            }

            throw (new ArgumentException("InterfaceNode not found " + refNode.Name));
        }

        public void ShowItems(XmlNode classNode)
        {

            XmlNode methodsNode = classNode.SelectSingleNode("Methods");
            methodsControlMain.ShowItems(methodsNode);

            XmlNode propertiesNode = classNode.SelectSingleNode("Properties");
            propertiesControlMain.ShowItems(propertiesNode);

            XmlNode eventsNode = classNode.SelectSingleNode("Interfaces/Events");
            eventsControlMain.Clear();
            foreach (XmlNode itemEventFaceNode in eventsNode.ChildNodes)
            {
                XmlNode rootEventInterfaceNode = GetEventInterfaceNode(itemEventFaceNode);
                XmlNode eventMethodsNode =  rootEventInterfaceNode.SelectSingleNode("Methods");

                eventsControlMain.ShowItems(eventMethodsNode, false);
            }
           
            ShowComponents(classNode);
        }

        #endregion

        #region Private Methods

        private void ShowComponents(XmlNode classNode)
        {
            dataGridViewCoClassComponents.Rows.Clear();
            XmlNode componentsNode = classNode.SelectSingleNode("Components");
            foreach (XmlNode componentNode in componentsNode.ChildNodes)
            {
                dataGridViewCoClassComponents.Rows.Add();
                DataGridViewRow newRow = dataGridViewCoClassComponents.Rows[dataGridViewCoClassComponents.Rows.Count - 1];
                XmlNode rootComponent = GetRootComponentNode(componentNode);
                newRow.Cells[0].Value = GetChildInnerText(rootComponent, "VersionAttribute");
                newRow.Cells[1].Value = GetChildInnerText(rootComponent, "Description");
                newRow.Cells[2].Value = GetChildInnerText(rootComponent, "ContainingFile");
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }

            dataGridViewCoClassInterfaces.Rows.Clear();
            XmlNode interfacesNode = classNode.SelectSingleNode("Interfaces/Implied");
            foreach (XmlNode itemNode in interfacesNode.ChildNodes)
            {
                dataGridViewCoClassInterfaces.Rows.Add();
                DataGridViewRow newRow = dataGridViewCoClassInterfaces.Rows[dataGridViewCoClassInterfaces.Rows.Count - 1];
                newRow.Cells[0].Value = "Implied";
                newRow.Cells[1].Value = itemNode.Attributes["Name"].InnerText;
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }

            interfacesNode = classNode.SelectSingleNode("Interfaces/VTable");
            foreach (XmlNode itemNode in interfacesNode.ChildNodes)
            {
                dataGridViewCoClassInterfaces.Rows.Add();
                DataGridViewRow newRow = dataGridViewCoClassInterfaces.Rows[dataGridViewCoClassInterfaces.Rows.Count - 1];
                newRow.Cells[0].Value = "VTable";
                newRow.Cells[1].Value = itemNode.Attributes["Name"].InnerText;
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }

            interfacesNode = classNode.SelectSingleNode("Interfaces/Events");
            foreach (XmlNode itemNode in interfacesNode.ChildNodes)
            {
                dataGridViewCoClassInterfaces.Rows.Add();
                DataGridViewRow newRow = dataGridViewCoClassInterfaces.Rows[dataGridViewCoClassInterfaces.Rows.Count - 1];
                newRow.Cells[0].Value = "Events";
                newRow.Cells[1].Value = itemNode.Attributes["Name"].InnerText;
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }

        }

        private void SetupCoClassInterfacesGrid()
        {
            dataGridViewCoClassInterfaces.Columns.Clear();
            dataGridViewCoClassInterfaces.Columns.Add("Type", "Type");
            dataGridViewCoClassInterfaces.Columns.Add("Name", "Name");
            dataGridViewCoClassInterfaces.Columns[0].ReadOnly = true;
            dataGridViewCoClassInterfaces.Columns[1].ReadOnly = true;
            dataGridViewCoClassInterfaces.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewCoClassInterfaces.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        }

        private void SetupCoClassComponentsGrid()
        {
            dataGridViewCoClassComponents.Columns.Clear();
            dataGridViewCoClassComponents.Columns.Add("Version", "Version");
            dataGridViewCoClassComponents.Columns.Add("Description", "Description");
            dataGridViewCoClassComponents.Columns.Add("ContainingFile", "ContainingFile");
            dataGridViewCoClassComponents.Columns[0].ReadOnly = true;
            dataGridViewCoClassComponents.Columns[1].ReadOnly = true;
            dataGridViewCoClassComponents.Columns[2].ReadOnly = true;

            dataGridViewCoClassComponents.Columns[0].Width = dataGridViewCoClassComponents.Width / 100 * 10;
            dataGridViewCoClassComponents.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCoClassComponents.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
 
        private XmlNode GetRootComponentNode(XmlNode refComponentNode)
        {
            string key = refComponentNode.Name;
            XmlNode rootComponents = refComponentNode.OwnerDocument.SelectSingleNode(Constants.Components);
            foreach (XmlNode rootComponentNode in rootComponents.ChildNodes)
            {
                XmlNode keyNode = GetChildNode(rootComponentNode, "Key");
                if (keyNode.InnerText == refComponentNode.Name)
                    return rootComponentNode;
            }
            throw (new ArgumentException("ComponentNode not found. " + refComponentNode.Name));
        }

        private XmlNode GetChildNode(XmlNode node, string name)
        {
            foreach (XmlNode item in node.ChildNodes)
            {
                if (item.Name == name)
                    return item;
            }
            throw (new ArgumentException("Node not found" + name));
        }

        private string GetChildInnerText(XmlNode node, string name)
        {
            foreach (XmlNode item in node.ChildNodes)
            {
                if (item.Name == name)
                    return item.InnerText;
            }
            throw (new ArgumentException("Node not found" + name));
        }

        #endregion     
    }
}
