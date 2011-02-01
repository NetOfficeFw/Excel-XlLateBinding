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
    public partial class MethodsGrid : UserControl
    {
        #region Fields

        ItemFilter _filter;
        bool _initializeFlag;
        
        #endregion

        #region Construction

        public MethodsGrid()
        {
            InitializeComponent();
        
            dataGridViewMethods.Columns.Add("Return", "Return");
            dataGridViewMethods.Columns.Add("Caption", "Caption");
            dataGridViewMethods.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells; 
            dataGridViewMethods.Columns.Add("Versions", "Versions");
            dataGridViewMethods.Columns.Add("Key", "Key");

            DataGridViewButtonColumn newButtonColumn = new DataGridViewButtonColumn();
            dataGridViewMethods.Columns.Add(newButtonColumn);
            newButtonColumn.HeaderText = "Delete";
           
            newButtonColumn = new DataGridViewButtonColumn();
            dataGridViewMethods.Columns.Add(newButtonColumn);
            newButtonColumn.HeaderText = "Edit";
        }
        
        #endregion

        #region Public Methods

        public void Clear()
        {
            dataGridViewMethods.Rows.Clear();
            parametersGridMain.Clear();
        }
           
        public void Show(ItemFilter filter, XmlNode node)
        {
            _filter = filter;

            _initializeFlag = true;
            this.Visible = true;
            Clear();
            if (node == null)
                return;
            
            dataGridViewMethods.Tag = node;

            XmlNode methodsNode = node.SelectSingleNode("Methods");
            foreach (XmlNode methodNode in methodsNode.ChildNodes)
            {
                string name =methodNode.Attributes["Name"].InnerText;
                foreach (XmlNode paramsNode in methodNode.SelectNodes("Parameters"))
                {
                    if(true == FilterPassed(paramsNode, name))
                    {
                        dataGridViewMethods.Rows.Add();
                        DataGridViewRow row = dataGridViewMethods.Rows[dataGridViewMethods.Rows.Count - 1];
                        row.Tag = methodNode;
                        row.Cells[0].Tag = paramsNode;
                        ShowMethod(row, methodNode, paramsNode);                       
                    }
                }
            }

            foreach (DataGridViewCell item in dataGridViewMethods.SelectedCells)
                item.Selected = false;
             
            _initializeFlag = false;
        }

        #endregion

        #region Private  Methods

        private string GetReturnValue(XmlNode paramsNode)
        {
            XmlNode returnValueNode = paramsNode.SelectSingleNode("ReturnValue");
            return returnValueNode.Attributes["Type"].InnerText;
        }
        
        private string GetVersions(XmlNode node)
        {
            string version = "";
            XmlNode refMemberComponents = node.SelectSingleNode("RefComponents");
            foreach (XmlNode refComponent in refMemberComponents.ChildNodes)
            {
                XmlNode componentNode = node.OwnerDocument.SelectSingleNode(XPathConstants.Components + "/Component[@Key='" + refComponent.Attributes["Key"].InnerText + "']");
                version += componentNode.Attributes["Version"].InnerText + " ";
            }
            return version.Trim();
        }

        private void ShowMethod(DataGridViewRow row, XmlNode methodNode, XmlNode paramsNode)
        {
            row.Cells[0].Value = GetReturnValue(paramsNode);
            row.Cells[1].Value = methodNode.Attributes["Caption"].InnerText;
            row.Cells[2].Value = GetVersions(paramsNode);
            row.Cells[3].Value = methodNode.Attributes["Key"].InnerText;
        }

        private void ShowParameter(DataGridViewRow row, XmlNode paramNode)
        {
            row.Cells[0].Value = paramNode.Attributes["Type"].InnerText;
            row.Cells[1].Value = paramNode.Attributes["Name"].InnerText;
            row.Cells[2].Value = paramNode.Attributes["IsOptional"].InnerText;
        }

        internal bool FilterPassed(XmlNode node, ItemFilter.FilterMode mode)
        {
            return _filter.FilterPassed(node, mode);
        }

        private bool FilterPassed(XmlNode node)
        {
            return _filter.FilterPassed(node);
        }

        private bool FilterPassed(XmlNode node, string expression)
        {
            bool globalFilterPassed = FilterPassed(node);
            if (false == globalFilterPassed)
                return false;

            string filterText = textBoxMethodFilter.Text.Trim();
            if (filterText == "") return true;

            if (expression.IndexOf(filterText, 0, StringComparison.InvariantCultureIgnoreCase) > -1)            
                return true;
            else
                return false;
        }

        #endregion

        #region Gui Trigger dataGridViewMethod

        private void dataGridViewMethods_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewMethods.Rows[e.RowIndex];

            XmlNode methodNode = (XmlNode)selectRow.Tag;
            XmlNode paramsNode = (XmlNode)selectRow.Cells[0].Tag;
            paramsNode.SelectSingleNode("ReturnValue").Attributes["Type"].InnerText = (string)selectRow.Cells[0].Value;
            methodNode.Attributes["Caption"].InnerText = (string)selectRow.Cells[1].Value;
        }

        private void dataGridViewMethods_SelectionChanged(object sender, EventArgs e)
        {
            if (true == _initializeFlag) return;
            if (dataGridViewMethods.SelectedCells.Count < 1) return;

            int rowIndex = dataGridViewMethods.SelectedCells[0].RowIndex;

            _initializeFlag = true;
            
            parametersGridMain.Clear();

            DataGridViewRow selectRow = dataGridViewMethods.Rows[rowIndex];
            DataGridViewCell selectCell = dataGridViewMethods.Rows[rowIndex].Cells[0];

            if (null != selectCell.Tag)
            {
                XmlNode paramsNode = (XmlNode)selectCell.Tag;
                parametersGridMain.Show(paramsNode);
            }

            _initializeFlag = false;
        }
 
        private void dataGridViewMethods_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if ((e.RowIndex < 0) || (true == _initializeFlag)) return;
            DataGridViewRow selectRow = dataGridViewMethods.Rows[e.RowIndex];
            if (dataGridViewMethods.Columns[e.ColumnIndex] is DataGridViewButtonColumn)
            {
                string typeAction = dataGridViewMethods.Columns[e.ColumnIndex].HeaderText;
                switch (typeAction)
                {
                    case "Edit":
                    { 
                        XmlNode methodNode = (XmlNode)selectRow.Tag;
                        XmlNode paramsNode = (XmlNode)selectRow.Cells[0].Tag;
                        FormEdit editForm = new FormEdit(methodNode);                        
                        DialogResult dr = editForm.ShowDialog(this);
                        if (dr == DialogResult.OK)
                        {
                            editForm.ApplyChanges(methodNode);
                            ShowMethod(selectRow, methodNode, paramsNode);
                        }
                        break;
                    }
                    case "Delete":
                    { 
                        string question = string.Format("Delete {0} ?", selectRow.Cells[1].Value);
                        DialogResult dr = MessageBox.Show(question, "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dr == DialogResult.Yes)
                        {
                            XmlNode methodNode = (XmlNode)selectRow.Tag;
                            XmlNode paramsNode = (XmlNode)selectRow.Cells[0].Tag;

                            methodNode.RemoveChild(paramsNode);
                            dataGridViewMethods.Rows.Remove(selectRow);

                            if (methodNode.SelectNodes("Parameters").Count == 0)
                            {
                                XmlNode methodsNode = methodNode.ParentNode;
                                methodsNode.RemoveChild(methodNode);
                            }

                        }
                    }
                    break;
                }
            }
        }
        
        private void textBoxMethodFilter_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Return)
            {
                XmlNode methodsNode = dataGridViewMethods.Tag as XmlNode;
                Show(_filter, methodsNode);
            }
        }

        private void buttonResolve_Click(object sender, EventArgs e)
        {
            XmlNode objectNode = dataGridViewMethods.Tag as XmlNode;
            FormResolveMethodConflict resolveForm = new FormResolveMethodConflict(new FilterPassedMethod(FilterPassed), objectNode,"Methods");
            DialogResult dr = resolveForm.ShowDialog(this);
            if (dr == DialogResult.OK)
            {
                XmlNode node = dataGridViewMethods.Tag as XmlNode;
                Show(_filter, node);
            }
        }

        #endregion     
    }
}
