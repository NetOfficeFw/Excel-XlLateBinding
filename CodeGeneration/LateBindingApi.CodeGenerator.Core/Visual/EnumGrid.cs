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
    public partial class EnumGrid : UserControl
    {
        bool _initializeFlag;

        public EnumGrid()
        {
            InitializeComponent();

            dataGridViewRefComponents.Columns.Add("Version", "Version");
            dataGridViewRefComponents.Columns.Add("Description", "Description");
            dataGridViewRefComponents.Columns.Add("Key", "Key");

            dataGridViewEnumMembers.Columns.Add("Name", "Name");
            dataGridViewEnumMembers.Columns.Add("Value", "Value");
            dataGridViewEnumMembers.Columns.Add("Versions", "Versions");

        }

        public void Clear()
        {
            dataGridViewRefComponents.Rows.Clear();
            dataGridViewEnumMembers.Rows.Clear();
        }

        public void Show(XmlNode node)
        {
            _initializeFlag = true;
            this.Visible = true;
            Clear();
            if (node == null)
                return;

            XmlNode refComponents = node.SelectSingleNode("RefComponents");
            foreach (XmlNode refComponent in refComponents.ChildNodes)
            {
                dataGridViewRefComponents.Rows.Add();
                DataGridViewRow row = dataGridViewRefComponents.Rows[dataGridViewRefComponents.Rows.Count - 1];

                XmlNode componentNode = node.OwnerDocument.SelectSingleNode(XPathConstants.Components + "/Component[@Key='" + refComponent.Attributes["Key"].InnerText + "']");

                row.Cells[0].Value = componentNode.Attributes["Version"].InnerText;
                row.Cells[1].Value = componentNode.Attributes["Description"].InnerText;
                row.Cells[2].Value = componentNode.Attributes["Key"].InnerText;
 
                row.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki); row.Cells[0].ReadOnly = true;
                row.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki); row.Cells[1].ReadOnly = true;
                row.Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki); row.Cells[2].ReadOnly = true;
                    
                row.Tag = componentNode;
            }
            
            XmlNode enumMembers = node.SelectSingleNode("EnumMembers");
            foreach (XmlNode memberNode in enumMembers.ChildNodes)
            {
                dataGridViewEnumMembers.Rows.Add();
                DataGridViewRow row = dataGridViewEnumMembers.Rows[dataGridViewEnumMembers.Rows.Count - 1];

                row.Cells[0].Value = memberNode.Attributes["Name"].InnerText;
                row.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki); row.Cells[0].ReadOnly = true;
                row.Cells[1].Value = memberNode.Attributes["Value"].InnerText;

                string version = "";
                XmlNode refMemberComponents = memberNode.SelectSingleNode("RefComponents");
                foreach (XmlNode refComponent in refMemberComponents.ChildNodes)
                {
                    XmlNode componentNode = node.OwnerDocument.SelectSingleNode(XPathConstants.Components + "/Component[@Key='" + refComponent.Attributes["Key"].InnerText + "']");
                    version += componentNode.Attributes["Version"].InnerText + " ";
                }
                row.Cells[2].Value = version.Trim();
                row.Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki); row.Cells[2].ReadOnly = true;

                row.Tag = memberNode;
            }

            _initializeFlag = false;
        }

        private void dataGridViewEnumMembers_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewEnumMembers.Rows[e.RowIndex];
            XmlNode enumMember = (XmlNode)selectRow.Tag;
            enumMember.Attributes["Value"].InnerText = (string)selectRow.Cells[1].Value;
        }
    }
}
