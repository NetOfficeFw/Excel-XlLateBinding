using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.Visual
{
    partial class FormEdit : Form
    {
        #region Fields
        
        XmlNode _itemNode;
        
        #endregion

        #region Properties

        public void ApplyChanges(XmlNode itemNode)
        {
            itemNode.InnerXml = ItemXml;
            foreach (DataGridViewRow itemRow in dataGridViewAttributes.Rows)
            {
                string name = (string)itemRow.Cells[0].Value;
                string val = (string)itemRow.Cells[1].Value;
                itemNode.Attributes[name].InnerText = val.Trim();
            }
        }

        public string ItemXml
        {
            get
            {
                return textBoxItem.Text.Trim();
            }
        }

        public DataGridView AttributesGrid
        {
            get
            {
                return dataGridViewAttributes;
            }   
        }

        #endregion

        #region Construction

        public FormEdit(XmlNode itemNode)
        {
            InitializeComponent();
            SetupAttributesGrid();
            _itemNode = itemNode;
            textBoxItem.Text = Utils.PrintXML(itemNode);
            
            dataGridViewAttributes.Rows.Clear();
            foreach (XmlAttribute attrib in itemNode.Attributes)
            {
                dataGridViewAttributes.Rows.Add();
                DataGridViewRow newRow = dataGridViewAttributes.Rows[dataGridViewAttributes.Rows.Count - 1];
                newRow.Cells[0].Value = attrib.Name;
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Value = attrib.InnerText;
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki);
            }
        }

        #endregion

        #region PrivateMethods

        private void SetupAttributesGrid()
        {
            dataGridViewAttributes.Columns.Add("Name", "Name");
            dataGridViewAttributes.Columns.Add("Value", "Value");
            dataGridViewAttributes.Columns[0].ReadOnly = true;
            dataGridViewAttributes.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewAttributes.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }
        
        #endregion

        #region Gui Trigger

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult =DialogResult.Cancel;
            this.Close();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        #endregion
    }
}
