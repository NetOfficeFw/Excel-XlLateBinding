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
    public partial class EnumControl : UserControl
    {
        #region Fields

        XmlNode _enumNode;

        #endregion
        
        #region Construction
        
        public EnumControl()
        {
            InitializeComponent();
            this.Visible = false;
            SetupGridColumns();
        }
        
        #endregion

        #region Private Methods

        private XmlNode GetChildNode(XmlNode node, string name)
        {
            foreach (XmlNode  item in node.ChildNodes)
            {
                if (item.Name == name)
                    return item;
            }
            throw(new ArgumentException("Node not found" + name));
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

        private XmlNode GetComponentNode(XmlNode node, string key)
        {
            XmlNode componentsNode = node.SelectSingleNode("Components");
            foreach (XmlNode childNode in componentsNode.ChildNodes)
            {
                XmlNode keyNode = GetChildNode(childNode, "Key");
                if (key == keyNode.InnerText)
                    return childNode;
            }
            
            throw (new ArgumentException("Node not found. " + key));
        }

        private string GetComponentDescriptions(XmlNode enumNode)
        {
            string result = "";
            XmlNode refComponents = GetChildNode(enumNode, "Components");
            foreach (XmlNode item in refComponents.ChildNodes)
            {
                string key = XmlConvert.DecodeName(item.Name);
                XmlNode componentNode = GetComponentNode(enumNode.OwnerDocument.FirstChild, key);
                result += GetChildInnerText(componentNode, "VersionAttribute") + "; ";
            }
            return result;
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
            dataGridViewComponents.Columns.Add("Description", "Description");
            dataGridViewComponents.Columns.Add("ContainingFile", "ContainingFile");

            dataGridViewComponents.Columns[1].Width = 320;
            dataGridViewComponents.Columns[2].Width = 360;
        }
        
        private void CheckDuplicateNameInEnumGrid()
        {
            foreach (DataGridViewRow row in dataGridViewEnum.Rows)
            {
                string name = (string)row.Cells[0].Value;
                int index = row.Index;
                for (int i = 0; i < dataGridViewEnum.Rows.Count; i++)
                {
                    if (i != index)
                    {
                        string otherName = (string)dataGridViewEnum.Rows[i].Cells[0].Value;
                        if (true == name.Equals(otherName, StringComparison.InvariantCultureIgnoreCase))
                            throw (new DuplicateNameException("name is duplicate " + name));
                    }
                }
            }
        }

        private void CheckIntegerValuesInEnumGrid()
        {
            foreach (DataGridViewRow row in dataGridViewEnum.Rows)
            {
                int val =0;
                string expression = (string)row.Cells[1].Value;
                bool sucseed = int.TryParse(expression, out val);
                if (false == sucseed)
                    throw (new FormatException(expression + " is not an integer value"));
            }
        }

        #endregion

        #region Public Methods

        public void Clear()
        {
        }


        public void ShowItem(XmlNode enumNode)
        {
            _enumNode = enumNode;
            dataGridViewComponents.Rows.Clear();
            dataGridViewEnum.Rows.Clear();

            XmlNode memberNode = GetChildNode(enumNode,"Members");
            XmlNode componentsNode = GetChildNode(enumNode, "Components");

            foreach (XmlNode componentNode in componentsNode.ChildNodes)
            {
                dataGridViewComponents.Rows.Add();
                int rowIndex = dataGridViewComponents.Rows.Count - 1;

                dataGridViewComponents.Rows[rowIndex].Tag = componentNode;
                string key = XmlConvert.DecodeName(componentNode.Name);
                XmlNode compNode = GetComponentNode(enumNode.ParentNode.ParentNode, key);
                dataGridViewComponents.Rows[rowIndex].Cells[0].Value = compNode.Name;
                dataGridViewComponents.Rows[rowIndex].Cells[1].Value = GetChildInnerText(compNode, "Description");
                dataGridViewComponents.Rows[rowIndex].Cells[2].Value = GetChildInnerText(compNode, "ContainingFile");

                dataGridViewComponents.Rows[rowIndex].Cells[0].ReadOnly = true;
                dataGridViewComponents.Rows[rowIndex].Cells[1].ReadOnly = true;
                dataGridViewComponents.Rows[rowIndex].Cells[2].ReadOnly = true;

                dataGridViewComponents.Rows[rowIndex].Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                dataGridViewComponents.Rows[rowIndex].Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                dataGridViewComponents.Rows[rowIndex].Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control);

            }

            foreach (XmlNode enumMemberNode in memberNode.ChildNodes)
            {
                dataGridViewEnum.Rows.Add();
                int rowIndex = dataGridViewEnum.Rows.Count - 1;

                dataGridViewEnum.Rows[rowIndex].Tag = enumMemberNode;

                dataGridViewEnum.Rows[rowIndex].Cells[0].Value = enumMemberNode.Name;
                dataGridViewEnum.Rows[rowIndex].Cells[1].Value = enumMemberNode.InnerText;
                dataGridViewEnum.Rows[rowIndex].Cells[2].Value = GetComponentDescriptions(enumMemberNode);

                dataGridViewEnum.Rows[rowIndex].Cells[0].ReadOnly = true;
                dataGridViewEnum.Rows[rowIndex].Cells[1].ReadOnly = true;
                dataGridViewEnum.Rows[rowIndex].Cells[2].ReadOnly = true;
                dataGridViewEnum.Rows[rowIndex].Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control); 
                dataGridViewEnum.Rows[rowIndex].Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control); 
                dataGridViewEnum.Rows[rowIndex].Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control); 

            }

        }
        
        #endregion

        #region Gui Trigger

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            try
            {
                ShowItem(_enumNode);
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }

        private XmlNode RenameNode(XmlNode node, string newName)
        {
            XmlNode newNode = node.OwnerDocument.CreateElement(newName);
            newNode.InnerXml = node.InnerXml;

            node.ParentNode.InsertBefore(newNode, node);
            node.ParentNode.RemoveChild(node);
            return newNode;
        }

        private void buttonApply_Click(object sender, EventArgs e)
        {
            try
            {
                CheckDuplicateNameInEnumGrid();
                CheckIntegerValuesInEnumGrid();

                foreach (DataGridViewRow row in dataGridViewEnum.Rows)
                {
                    string name = (string)row.Cells[0].Value;
                  
                    XmlNode enumMembersNode = GetChildNode(_enumNode,"Members");
                    XmlNode enumMemberNode = GetChildNode(enumMembersNode, name);

                    enumMemberNode.ChildNodes[0].InnerText   = (string)row.Cells[1].Value;
                }

            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }

        private void dataGridViewEnum_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }
        
        #endregion
    }
}
