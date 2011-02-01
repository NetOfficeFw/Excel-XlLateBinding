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
    public partial class ParametersGrid : UserControl
    {     
        #region Fields

        bool _initializeFlag;
        
        #endregion

        public ParametersGrid()
        {
            InitializeComponent();

            dataGridViewMethodParameters.Columns.Add("Type", "Type");
            dataGridViewMethodParameters.Columns.Add("Name", "Name");
            dataGridViewMethodParameters.Columns.Add("IsOptional", "IsOptional");

            DataGridViewButtonColumn newButtonColumn = new DataGridViewButtonColumn();
            dataGridViewMethodParameters.Columns.Add(newButtonColumn);
            newButtonColumn.HeaderText = "Delete";

            newButtonColumn = new DataGridViewButtonColumn();
            dataGridViewMethodParameters.Columns.Add(newButtonColumn);
            newButtonColumn.HeaderText = "Edit";
        }

        #region Public Methods

        public void Clear()
        {
            dataGridViewMethodParameters.Rows.Clear();
        }

        public void Show(XmlNode paramsNode)
        {
            _initializeFlag = true;
            Clear();
            if (paramsNode == null)
                return;

            foreach (XmlNode paramNode in paramsNode.SelectNodes("Parameter"))
            {
                dataGridViewMethodParameters.Rows.Add();
                DataGridViewRow newRow = dataGridViewMethodParameters.Rows[dataGridViewMethodParameters.Rows.Count - 1];
                ShowParameter(newRow, paramNode);
                newRow.Tag = paramNode;
            }

            _initializeFlag = false;
        }

        #endregion

        #region Private Methods

        private void ShowParameter(DataGridViewRow row, XmlNode paramNode)
        {
            row.Cells[0].Value = paramNode.Attributes["Type"].InnerText;
            row.Cells[1].Value = paramNode.Attributes["Name"].InnerText;
            row.Cells[2].Value = paramNode.Attributes["IsOptional"].InnerText;
        }

        #endregion

        #region Gui Trigger dataGridViewMethodParameters

        private void dataGridViewMethodParameters_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewMethodParameters.Rows[e.RowIndex];
            XmlNode paramNode = (XmlNode)selectRow.Tag;
            paramNode.Attributes["Type"].InnerText = (string)selectRow.Cells[0].Value;
            paramNode.Attributes["Name"].InnerText = (string)selectRow.Cells[1].Value;
        }

        private void dataGridViewMethodParameters_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewMethodParameters.Rows[e.RowIndex];
            if (dataGridViewMethodParameters.Columns[e.ColumnIndex] is DataGridViewButtonColumn)
            {
                string typeAction = dataGridViewMethodParameters.Columns[e.ColumnIndex].HeaderText;
                switch (typeAction)
                {
                    case "Edit":
                        {
                            XmlNode paramNode = (XmlNode)selectRow.Tag;
                            FormEdit editForm = new FormEdit(paramNode);
                            DialogResult dr = editForm.ShowDialog(this);
                            if (dr == DialogResult.OK)
                            {
                                editForm.ApplyChanges(paramNode);
                                ShowParameter(selectRow, paramNode);
                            }
                            break;
                        }
                    case "Delete":
                        {
                            string question = string.Format("Delete {0} ?", selectRow.Cells[1].Value);
                            DialogResult dr = MessageBox.Show(question, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (dr == DialogResult.Yes)
                            {
                                XmlNode parameterNode = (XmlNode)selectRow.Tag;
                                XmlNode parametersNode = parameterNode.ParentNode;
                                parametersNode.RemoveChild(parameterNode);
                                dataGridViewMethodParameters.Rows.Remove(selectRow);
                            }
                            break;
                        }
                }
            }
        }

        #endregion
    }
}
