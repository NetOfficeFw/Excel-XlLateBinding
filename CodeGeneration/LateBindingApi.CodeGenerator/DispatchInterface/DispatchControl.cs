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
    public partial class DispatchControl : UserControl
    {
        public DispatchControl()
        {
            InitializeComponent();
            this.Visible = false;
            methodsControlMain.Initialize();
            propertiesControlMain.Initialize();
            SetupInterfaceComponentsGrid();
            SetupInterfaceInterfacesGrid();
        }

        public void ShowItems(XmlNode interfaceNode)
        {
            XmlNode methodsNode = interfaceNode.SelectSingleNode("Methods");
            methodsControlMain.ShowItems(methodsNode);

            XmlNode propertiesNode = interfaceNode.SelectSingleNode("Properties");
            propertiesControlMain.ShowItems(propertiesNode);

            ShowComponents(interfaceNode);
        }

        public void Clear()
        {
            dataGridViewInterfaceInterfaces.Rows.Clear();
            dataGridViewInterfaceComponents.Rows.Clear();
            methodsControlMain.Clear();
            propertiesControlMain.Clear();
        }

        #region Private Methods

        private void ShowComponents(XmlNode interfaceNode)
        {
            dataGridViewInterfaceComponents.Rows.Clear();
            XmlNode componentsNode = interfaceNode.SelectSingleNode("Components");
            foreach (XmlNode componentNode in componentsNode.ChildNodes)
            {
                dataGridViewInterfaceComponents.Rows.Add();
                DataGridViewRow newRow = dataGridViewInterfaceComponents.Rows[dataGridViewInterfaceComponents.Rows.Count - 1];
                XmlNode rootComponent = GetRootComponentNode(componentNode);
                newRow.Cells[0].Value = GetChildInnerText(rootComponent, "VersionAttribute");
                newRow.Cells[1].Value = GetChildInnerText(rootComponent, "Description");
                newRow.Cells[2].Value = GetChildInnerText(rootComponent, "ContainingFile");
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }

            dataGridViewInterfaceInterfaces.Rows.Clear();
            XmlNode interfacesNode = interfaceNode.SelectSingleNode("Interfaces/Implied");
            foreach (XmlNode itemNode in interfacesNode.ChildNodes)
            {
                dataGridViewInterfaceInterfaces.Rows.Add();
                DataGridViewRow newRow = dataGridViewInterfaceInterfaces.Rows[dataGridViewInterfaceInterfaces.Rows.Count - 1];
                newRow.Cells[0].Value = itemNode.Attributes["Type"].InnerText;
                newRow.Cells[1].Value = itemNode.Attributes["Name"].InnerText;
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }

            interfacesNode = interfaceNode.SelectSingleNode("Interfaces/VTable");
            foreach (XmlNode itemNode in interfacesNode.ChildNodes)
            {
                dataGridViewInterfaceInterfaces.Rows.Add();
                DataGridViewRow newRow = dataGridViewInterfaceInterfaces.Rows[dataGridViewInterfaceInterfaces.Rows.Count - 1];
                newRow.Cells[0].Value = itemNode.Attributes["Type"].InnerText;
                newRow.Cells[1].Value = itemNode.Attributes["Name"].InnerText;
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }
        }

        private void SetupInterfaceInterfacesGrid()
        {
            dataGridViewInterfaceInterfaces.Columns.Clear();
            dataGridViewInterfaceInterfaces.Columns.Add("Type", "Type");
            dataGridViewInterfaceInterfaces.Columns.Add("Name", "Name");
            dataGridViewInterfaceInterfaces.Columns[0].ReadOnly = true;
            dataGridViewInterfaceInterfaces.Columns[1].ReadOnly = true;
            dataGridViewInterfaceInterfaces.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dataGridViewInterfaceInterfaces.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
        }

        private void SetupInterfaceComponentsGrid()
        {
            dataGridViewInterfaceComponents.Columns.Clear();
            dataGridViewInterfaceComponents.Columns.Add("Version", "Version");
            dataGridViewInterfaceComponents.Columns.Add("Description", "Description");
            dataGridViewInterfaceComponents.Columns.Add("ContainingFile", "ContainingFile");
            dataGridViewInterfaceComponents.Columns[0].ReadOnly = true;
            dataGridViewInterfaceComponents.Columns[1].ReadOnly = true;
            dataGridViewInterfaceComponents.Columns[2].ReadOnly = true;

            dataGridViewInterfaceComponents.Columns[0].Width = dataGridViewInterfaceComponents.Width / 100 * 10;
            dataGridViewInterfaceComponents.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewInterfaceComponents.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
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
