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
    public partial class RefInterfacesGrid : UserControl
    {
     
        public RefInterfacesGrid()
        {
            InitializeComponent();

            dataGridViewRefInterfaces.Columns.Add("Name", "Name");
            dataGridViewRefInterfaces.Columns.Add("Key", "Key");
        }

        public void Clear()
        {
            dataGridViewRefInterfaces.Rows.Clear();
        }

        public void Show(XmlNode node, string target)
        {
            Clear();
            if (node == null)
                return;

            XmlNodeList refInterfaces = node.SelectNodes(target);
            foreach (XmlNode refInterface in refInterfaces)
            {
                XmlNode component = GetInterfaceNode(refInterface);
                dataGridViewRefInterfaces.Rows.Add();
                DataGridViewRow row = dataGridViewRefInterfaces.Rows[dataGridViewRefInterfaces.Rows.Count - 1];
                WriteReadOnlyCell(row.Cells[0], component.Attributes["Name"].InnerText);
                WriteReadOnlyCell(row.Cells[1], component.Attributes["Key"].InnerText);
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

    }
}
