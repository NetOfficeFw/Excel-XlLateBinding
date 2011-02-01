using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core
{
    public partial class EnumControl : UserControl
    {
        #region Fields

        bool _initializeFlag;

        #endregion

        #region Construction

        public EnumControl()
        {
            InitializeComponent();
            this.Visible = false;
            SetupGridColumns();
        }
        #endregion
    
        #region Public Methods

        public void ShowItems(XmlNode itemNode)
        {
            _initializeFlag = true;

            dataGridViewComponents.Rows.Clear();
            dataGridViewEnum.Rows.Clear();

            XmlNode memberNode = itemNode.SelectSingleNode("Members");
            XmlNode componentsNode = itemNode.SelectSingleNode("Components");

            foreach (XmlNode componentNode in componentsNode.ChildNodes)
            {
                dataGridViewComponents.Rows.Add();
                int rowIndex = dataGridViewComponents.Rows.Count - 1;

                dataGridViewComponents.Rows[rowIndex].Tag = componentNode;
                string key = componentNode.InnerText;
                XmlNode compNode = GetComponentNode(itemNode.OwnerDocument, key);
                dataGridViewComponents.Rows[rowIndex].Cells[0].Value = compNode.Attributes["Name"].InnerText;
                dataGridViewComponents.Rows[rowIndex].Cells[1].Value = GetChildInnerText(compNode, "VersionAttribute");
                dataGridViewComponents.Rows[rowIndex].Cells[2].Value = GetChildInnerText(compNode, "Description");
                dataGridViewComponents.Rows[rowIndex].Cells[3].Value = GetChildInnerText(compNode, "ContainingFile");

                dataGridViewComponents.Rows[rowIndex].Cells[0].ReadOnly = true;
                dataGridViewComponents.Rows[rowIndex].Cells[1].ReadOnly = true;
                dataGridViewComponents.Rows[rowIndex].Cells[2].ReadOnly = true;
                dataGridViewComponents.Rows[rowIndex].Cells[3].ReadOnly = true;

                dataGridViewComponents.Rows[rowIndex].Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                dataGridViewComponents.Rows[rowIndex].Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                dataGridViewComponents.Rows[rowIndex].Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                dataGridViewComponents.Rows[rowIndex].Cells[3].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
            }

            foreach (XmlNode enumMemberNode in memberNode.ChildNodes)
            {
                dataGridViewEnum.Rows.Add();
                int rowIndex = dataGridViewEnum.Rows.Count - 1;

                dataGridViewEnum.Rows[rowIndex].Tag = enumMemberNode;

                dataGridViewEnum.Rows[rowIndex].Cells[0].Value = enumMemberNode.Attributes["Name"].InnerText;
                dataGridViewEnum.Rows[rowIndex].Cells[1].Value = enumMemberNode.Attributes["Value"].InnerText;
                dataGridViewEnum.Rows[rowIndex].Cells[2].Value = GetComponentDescriptions(enumMemberNode);

                dataGridViewEnum.Rows[rowIndex].Cells[0].ReadOnly = true;
                dataGridViewEnum.Rows[rowIndex].Cells[1].ReadOnly = true;
                dataGridViewEnum.Rows[rowIndex].Cells[2].ReadOnly = true;
                dataGridViewEnum.Rows[rowIndex].Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                dataGridViewEnum.Rows[rowIndex].Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                dataGridViewEnum.Rows[rowIndex].Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control);

            }

            _initializeFlag = false;
        }

        public void Clear()
        {
            _initializeFlag = true;
            dataGridViewComponents.Rows.Clear();
            dataGridViewEnum.Rows.Clear();           
            _initializeFlag = false;
        }

        #endregion

        #region Private Methods

        private string GetComponentDescriptions(XmlNode enumNode)
        {
            string result = "";

            XmlNode refComponents = enumNode.SelectSingleNode("Components");
            foreach (XmlNode item in refComponents.ChildNodes)
            {
                string key = item.Name;
                XmlNode componentNode = GetComponentNode(enumNode.OwnerDocument, key);
                result += GetChildInnerText(componentNode, "VersionAttribute") + "; ";
            }
            return result;
        }

        private XmlNode GetComponentNode(XmlDocument ownerDocument, string key)
        {
            XmlNode componentsNode = ownerDocument.SelectSingleNode(Constants.Components);
            foreach (XmlNode childNode in componentsNode.ChildNodes)
            {
                XmlNode keyNode = childNode.SelectSingleNode("Key");
                XmlNode extNode = childNode.SelectSingleNode("ExternalDependency");
                if( (key == keyNode.InnerText) &&  ("False" == extNode.InnerText) )
                    return childNode;

            }
            throw (new ArgumentException("Node not found. " + key));
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

        private void SetupGridColumns()
        {
            dataGridViewEnum.Columns.Add("Field", "Field");
            dataGridViewEnum.Columns[dataGridViewEnum.Columns.Count - 1].Width = 180;
            dataGridViewEnum.Columns.Add("Value", "Value");
            dataGridViewEnum.Columns.Add("Components", "Components");

            dataGridViewEnum.Columns[dataGridViewEnum.Columns.Count - 1].Width = 500;
            dataGridViewEnum.Columns[dataGridViewEnum.Columns.Count - 1].ReadOnly = true;

            dataGridViewComponents.Columns.Add("Name", "Name");
            dataGridViewComponents.Columns.Add("Version", "Version");
            dataGridViewComponents.Columns.Add("Description", "Description");
            dataGridViewComponents.Columns.Add("ContainingFile", "ContainingFile");

            dataGridViewComponents.Columns[2].Width = 320;
            dataGridViewComponents.Columns[3].Width = 360;
        }

        #endregion
    }
}
