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
    public partial class ChildAttributesGrid : UserControl
    {
        bool _initializeFlag;

        public ChildAttributesGrid()
        {
            InitializeComponent();
        }

        public void Clear()
        {
            dataGridViewAttributes.Columns.Clear();
            dataGridViewAttributes.Rows.Clear();
        }

        public void Show(XmlNode node)
        {
            _initializeFlag = true;            
            this.Visible = true;
            Clear();
            if (node == null)
                return;

            if( node.ChildNodes.Count == 0)
               return;
             
            XmlAttributeCollection templateAttribute = node.ChildNodes[0].Attributes;
            foreach (XmlAttribute attribute in templateAttribute)
            {
                dataGridViewAttributes.Columns.Add(attribute.Name, attribute.Name);
            }

            foreach (XmlNode nodeChild in node.ChildNodes)
            {
                dataGridViewAttributes.Rows.Add();
                DataGridViewRow row = dataGridViewAttributes.Rows[dataGridViewAttributes.Rows.Count - 1];
                row.Tag = nodeChild;
 
                int i = 0;
                foreach (XmlAttribute itemAttribute in nodeChild.Attributes)
                {
                    row.Cells[i].Value = itemAttribute.InnerText;
                    row.Cells[i].Tag = itemAttribute;
                    i++;
                }                
            }

            _initializeFlag = false;
        }

        private void dataGridViewAttributes_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewCell selectCell = dataGridViewAttributes.Rows[e.RowIndex].Cells[e.ColumnIndex];
            XmlAttribute attribute = (XmlAttribute)selectCell.Tag;
            attribute.InnerText = (string)selectCell.Value;
        }

    }
}
