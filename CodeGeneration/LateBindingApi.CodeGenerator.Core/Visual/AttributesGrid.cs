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
    public partial class AttributesGrid : UserControl
    {
        bool _initializeFlag;

        public AttributesGrid()
        {
            InitializeComponent();
            dataGridViewAttributes.Columns.Add("Name", "Name");
            dataGridViewAttributes.Columns.Add("Value", "Value");
            dataGridViewAttributes.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; 
            dataGridViewAttributes.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        public void Clear()
        {
            dataGridViewAttributes.Rows.Clear();
        }

        public void Show(XmlNode node)
        {
            _initializeFlag = true;
            this.Visible = true;
            Clear();
            if (node == null)
                return;

            foreach (XmlAttribute itemAttribute in node.Attributes)
            {
                dataGridViewAttributes.Rows.Add();
                DataGridViewRow row = dataGridViewAttributes.Rows[dataGridViewAttributes.Rows.Count - 1];
                row.Cells[0].Value = itemAttribute.Name;
                row.Cells[1].Value = itemAttribute.InnerText;

                row.Cells[0].ReadOnly = true;
                row.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                row.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki);

                row.Tag = itemAttribute;
            }

            _initializeFlag = false;
        }

        private void dataGridViewAttributes_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewAttributes.Rows[e.RowIndex];
            XmlAttribute attribute = (XmlAttribute)selectRow.Tag;
            attribute.InnerText = (string)selectRow.Cells[1].Value;
        }
    }
}
