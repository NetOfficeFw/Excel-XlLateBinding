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
    public partial class RefComponentsGrid : UserControl
    {
        public RefComponentsGrid()
        {
            InitializeComponent();

            dataGridViewComponents.Columns.Add("Name", "Name");
            dataGridViewComponents.Columns.Add("Version", "Version");
            dataGridViewComponents.Columns.Add("IsExternal", "IsExternal");
            dataGridViewComponents.Columns.Add("Key", "Key");
        }

        public void Clear()
        {
            dataGridViewComponents.Rows.Clear();
        }
       
        public void Show(XmlNode node)
        {
            Clear();
            if (node == null)
                return;

            XmlNodeList refComponents = node.SelectNodes("RefComponents/RefComponent");
            foreach (XmlNode refComponent in refComponents)
            {
                XmlNode component = GetComponentNode(refComponent);
                dataGridViewComponents.Rows.Add();
                DataGridViewRow row = dataGridViewComponents.Rows[dataGridViewComponents.Rows.Count - 1];
               
                string isExternal = "False";
                if (component.Name == "External")
                    isExternal = "True";

                WriteReadOnlyCell(row.Cells[0], component.Attributes["Name"].InnerText);
                WriteReadOnlyCell(row.Cells[1], component.Attributes["Version"].InnerText);
                WriteReadOnlyCell(row.Cells[2], isExternal);
                WriteReadOnlyCell(row.Cells[3], component.Attributes["Key"].InnerText);
            } 
        }

        private void WriteReadOnlyCell(DataGridViewCell cell, string value)
        {
            cell.Value = value;
            cell.ReadOnly = true;
            cell.Style.BackColor = Color.FromKnownColor(KnownColor.Control);
        }

        private void WriteCell(DataGridViewCell cell, string value)
        {
            cell.Value = value;
            cell.Style.BackColor = Color.DarkKhaki;
        }

        private XmlNode GetComponentNode(XmlNode refComponent)
        {
            string key = refComponent.Attributes["Key"].InnerText;
            XmlNode node = refComponent.OwnerDocument.SelectSingleNode(XPathConstants.Components + "/Component[@Key ='" + key + "']");
            if(node != null)
                return node;

            node = refComponent.OwnerDocument.SelectSingleNode(XPathConstants.Externals + "/External[@Key ='" + key + "']");
            if (node != null)
                return node;

            throw (new ArgumentException("Component not found" + key));           
        }
    }
}
