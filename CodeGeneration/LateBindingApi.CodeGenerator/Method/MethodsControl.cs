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
    public partial class MethodsControl : UserControl
    {
        #region Fields

        bool _initializeFlag;
        XmlNode _methodsNode;

        #endregion

        #region Construction

        public MethodsControl()
        {
            InitializeComponent();
            Initialize();
        }

        #endregion

        #region Public Methods

        public void Clear()
        {
            dataGridViewMethods.Rows.Clear();
            dataGridViewMethodParams.Rows.Clear();
        }

        public void Initialize()
        {
            SetupGridColumns();
        }

        public void ShowItems(XmlNode methodsNode)
        {
            _methodsNode = methodsNode;
            _initializeFlag = true;

            dataGridViewMethods.Rows.Clear();
            dataGridViewMethodParams.Rows.Clear();

            foreach (XmlNode methodNode in methodsNode.ChildNodes)
            {
                if (methodNode.Attributes["Name"].InnerText.IndexOf(textBoxMethodFilter.Text, StringComparison.InvariantCultureIgnoreCase)> -1)
                { 
                    dataGridViewMethods.Rows.Add();
                    DataGridViewRow newRow = dataGridViewMethods.Rows[dataGridViewMethods.Rows.Count - 1];
                    newRow.Cells[0].Value = methodNode.Attributes["ReturnType"].InnerText;
                    newRow.Cells[1].Value = methodNode.Attributes["Name"].InnerText;
                    newRow.Cells[2].Value = GetComponentVersionAttributes(methodNode.SelectSingleNode("Components"));
                    newRow.Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                    newRow.Tag = methodNode;
                }
            }

            _initializeFlag = false;
        }
        
        #endregion

        #region Private Methods

        private void SetupGridColumns()
        {
            dataGridViewMethods.Columns.Clear();
            dataGridViewMethods.Columns.Add("ReturnType", "ReturnType");
            dataGridViewMethods.Columns.Add("Name", "Name");
            dataGridViewMethods.Columns.Add("Components", "Components");
            dataGridViewMethods.Columns[2].ReadOnly = true;
            DataGridViewButtonColumn newButtonColumn = new DataGridViewButtonColumn();
            dataGridViewMethods.Columns.Add(newButtonColumn);
            newButtonColumn.HeaderText = "Delete";
            
            newButtonColumn = new DataGridViewButtonColumn();
            dataGridViewMethods.Columns.Add(newButtonColumn);
            newButtonColumn.HeaderText = "Edit";

            dataGridViewMethodParams.Columns.Clear();
            dataGridViewMethodParams.Columns.Add("Type", "Type");
            dataGridViewMethodParams.Columns.Add("Name", "Name");
            dataGridViewMethodParams.Columns.Add("Optional", "Optional");
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

        private string GetComponentVersionAttributes(XmlNode refComponentsNode)
        {
            string result = "";
            foreach (XmlNode refNode in refComponentsNode.ChildNodes)
            {
                XmlNode rootComponentNode = GetRootComponentNode(refNode);
                result += GetChildInnerText(rootComponentNode, "VersionAttribute") + ";";
            }
            return result;
        }

        #endregion

        #region Gui Trigger

        private void dataGridViewMethods_SelectionChanged(object sender, EventArgs e)
        {
            if (true == _initializeFlag) return;
            if (dataGridViewMethods.SelectedCells.Count < 1) return;

            _initializeFlag = true;

            dataGridViewMethodParams.Rows.Clear();
            DataGridViewCell selectCell = dataGridViewMethods.SelectedCells[0];
            DataGridViewRow selectRow = dataGridViewMethods.Rows[selectCell.RowIndex];
            XmlNode methodNode = (XmlNode)selectRow.Tag;
            XmlNode paramsNode = methodNode.SelectSingleNode("Parameters");
            foreach (XmlNode paramNode in paramsNode.ChildNodes)
            {
                dataGridViewMethodParams.Rows.Add();
                DataGridViewRow newRow = dataGridViewMethodParams.Rows[dataGridViewMethodParams.Rows.Count - 1];
                newRow.Cells[0].Value = paramNode.Attributes["Type"].InnerText;
                newRow.Cells[1].Value = paramNode.Attributes["Name"].InnerText;
                newRow.Cells[2].Value = paramNode.Attributes["Optional"].InnerText;
                newRow.Tag = paramNode;
            }

            _initializeFlag = false;
        }

        private void dataGridViewMethods_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
                DataGridViewRow selectRow = dataGridViewMethods.Rows[e.RowIndex];
                if (e.ColumnIndex == 3)
                {
                    string question = string.Format("Delete {0} ?", selectRow.Cells[1].Value);
                    DialogResult dr = MessageBox.Show(question, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dr == DialogResult.Yes)
                    {
                        XmlNode methodNode = (XmlNode)selectRow.Tag;
                        XmlNode methodsNode = methodNode.ParentNode;
                        methodsNode.RemoveChild(methodNode);
                        dataGridViewMethods.Rows.Remove(selectRow);
                    }
                }
                else if (e.ColumnIndex == 4)
                {
                    XmlNode itemNode = (XmlNode)selectRow.Tag;

                    FormEditXmlItem editForm = new FormEditXmlItem(itemNode);
                    DialogResult dr = editForm.ShowDialog(this);
                    if (dr == DialogResult.OK)
                    {
                        editForm.ApplyChanges(itemNode);
                        selectRow.Cells[0].Value = itemNode.Attributes["ReturnType"].Value;
                        selectRow.Cells[1].Value = itemNode.Attributes["Name"].Value;
                       
                    }
                }

            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException, "Failed to load Project Information.");
                errorForm.ShowDialog(this);
            }
        }

        private void dataGridViewMethods_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewMethods.Rows[e.RowIndex];

            XmlNode methodNode = (XmlNode)selectRow.Tag;
            methodNode.Attributes["ReturnType"].InnerText = (string)selectRow.Cells[0].Value;
            methodNode.Attributes["Name"].InnerText = (string)selectRow.Cells[1].Value;
        }

        private void dataGridViewMethodParams_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewMethodParams.Rows[e.RowIndex];
            XmlNode paramNodeNode = (XmlNode)selectRow.Tag;
            paramNodeNode.Attributes["Type"].InnerText = (string)selectRow.Cells[0].Value;
            paramNodeNode.Attributes["Name"].InnerText = (string)selectRow.Cells[1].Value;
        }

        private void textBoxMethodFilter_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Return)
            {
                ShowItems(_methodsNode);
            }
        }

        #endregion
    }
}
