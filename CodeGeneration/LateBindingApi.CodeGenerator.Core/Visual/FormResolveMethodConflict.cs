using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.Visual
{
    public delegate bool FilterPassedMethod(XmlNode node, ItemFilter.FilterMode mode);

    public partial class FormResolveMethodConflict : Form
    {
        bool _initializeFlag;

        FilterPassedMethod _filterPassed;
        XmlNode     _objectNode;
        string      _itemKind;
        XmlNodeList listCoClasses;
        XmlNodeList listInterfaces;
        XmlNodeList listDispatchInterfaces;

        public FormResolveMethodConflict(FilterPassedMethod filterPassed, XmlNode objectNode, string itemKind)
        {
            InitializeComponent();
            _itemKind = itemKind;
            _filterPassed = filterPassed;
            _objectNode = objectNode;

            listCoClasses = _objectNode.OwnerDocument.SelectNodes(XPathConstants.Solution + "/Projects/Project/CoClasses/CoClass");
            listInterfaces = _objectNode.OwnerDocument.SelectNodes(XPathConstants.Solution + "/Projects/Project/DispatchInterfaces/Interface");
            listDispatchInterfaces = _objectNode.OwnerDocument.SelectNodes(XPathConstants.Solution + "/Projects/Project/Interfaces/Interface");

            dataGridViewMethods.Columns.Add("Return", "Return");
            dataGridViewMethods.Columns.Add("Caption", "Caption");
            dataGridViewMethods.Columns.Add("Param Count", "Param Count");

            DataGridViewCheckBoxColumn newCheckBoxColumn = new DataGridViewCheckBoxColumn();
            dataGridViewMethods.Columns.Add(newCheckBoxColumn);
            newCheckBoxColumn.HeaderText = "Resolve";

            DataGridViewComboBoxColumn comboColumn = new DataGridViewComboBoxColumn();
            dataGridViewMethods.Columns.Add(comboColumn);
            comboColumn.HeaderText = "Return Value Suggestions";

            dataGridViewMethodParameters.Columns.Add("Type", "Type");
            dataGridViewMethodParameters.Columns.Add("Caption", "Caption");
            dataGridViewMethodParameters.Columns.Add("IsOptional", "IsOptional");

            FillForm();            
        }
        
        private void FillForm()
        {
            dataGridViewMethods.Rows.Clear();
            dataGridViewMethodParameters.Rows.Clear();

            XmlNode methodsNode = _objectNode.SelectSingleNode(_itemKind);
            foreach (XmlNode methodNode in methodsNode.ChildNodes)
            {
                string name = methodNode.Attributes["Name"].InnerText;
                foreach (XmlNode paramsNode in methodNode.SelectNodes("Parameters"))
                {
                    ItemFilter.FilterMode mode = ItemFilter.FilterMode.All;
                    if (radioButtonConflicts.Checked == true)
                        mode = ItemFilter.FilterMode.Conflict;

                    if (true == _filterPassed(paramsNode, mode))
                    {
                        dataGridViewMethods.Rows.Add();
                        DataGridViewRow row = dataGridViewMethods.Rows[dataGridViewMethods.Rows.Count - 1];
                        row.Tag = methodNode;
                        row.Cells[0].Tag = paramsNode;

                        WriteReadOnlyCell(row.Cells[0], GetReturnValue(paramsNode));
                        WriteReadOnlyCell(row.Cells[1], methodNode.Attributes["Caption"].InnerText);

                        int paramCount = 0;
                        if (null != paramsNode.SelectNodes("Parameter"))
                            paramCount = paramsNode.SelectNodes("Parameter").Count;

                        WriteReadOnlyCell(row.Cells[2], paramCount.ToString());

                        DataGridViewComboBoxCell comboCell = row.Cells[4] as DataGridViewComboBoxCell;
                        FillSuggestions(comboCell, methodNode.Attributes["Caption"].InnerText);
                    }
                }
            }

            _initializeFlag = false;
        }

        private string GetReturnValue(XmlNode methodNode)
        {
            XmlNode returnValueNode = methodNode.SelectSingleNode("ReturnValue");
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

        public void FillSuggestions(DataGridViewComboBoxCell comboCell, string name)
        {
            List<XmlNode> suggestionsList = new List<XmlNode>();
            foreach (XmlNode coClassNode in listCoClasses)
            {
                string caption = coClassNode.Attributes["Caption"].InnerText;
                if (name.IndexOf(caption, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    comboCell.Items.Add(caption);
                    suggestionsList.Add(coClassNode);
                }
            }

            foreach (XmlNode coClassNode in listDispatchInterfaces)
            {
                string caption = coClassNode.Attributes["Caption"].InnerText;
                if (name.IndexOf(caption, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    comboCell.Items.Add(caption);
                    suggestionsList.Add(coClassNode);
                }
            }

            foreach (XmlNode coClassNode in listInterfaces)
            {
                string caption = coClassNode.Attributes["Caption"].InnerText;
                if (name.IndexOf(caption, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    comboCell.Items.Add(caption);
                    suggestionsList.Add(coClassNode);
                }
            }

            comboCell.Items.Add("");
            if (suggestionsList.Count > 0)
                comboCell.Value = suggestionsList[0].Attributes["Caption"].InnerText;

            FillScalarTypes(comboCell);
            comboCell.Items.Add("");
            foreach (XmlNode coClassNode in listCoClasses)
            {
                string caption = coClassNode.Attributes["Caption"].InnerText;
                comboCell.Items.Add(caption);
                suggestionsList.Add(coClassNode);
            }
            foreach (XmlNode coClassNode in listDispatchInterfaces)
            {
                string caption = coClassNode.Attributes["Caption"].InnerText;
                comboCell.Items.Add(caption);
                suggestionsList.Add(coClassNode);
            }
            foreach (XmlNode coClassNode in listInterfaces)
            {
                string caption = coClassNode.Attributes["Caption"].InnerText;
                comboCell.Items.Add(caption);
                suggestionsList.Add(coClassNode);
            }

            comboCell.Tag = suggestionsList;
        }

        private void FillScalarTypes(DataGridViewComboBoxCell comboCell)
        {
            comboCell.Items.Add("bool");
            comboCell.Items.Add("string");
            comboCell.Items.Add("Double");
            comboCell.Items.Add("Int32");
        }

        private bool IsScalar(string expression)
        {
            switch (expression)
            {
                case "bool":
                case "string":
                case "Double":
                case "Int32":
                    return true;
            }
            return false;
        }

        private void ShowParameter(DataGridViewRow row, XmlNode paramNode)
        {
            WriteReadOnlyCell(row.Cells[0], paramNode.Attributes["Type"].InnerText);
            WriteReadOnlyCell(row.Cells[1], paramNode.Attributes["Name"].InnerText);
            WriteReadOnlyCell(row.Cells[2], paramNode.Attributes["IsOptional"].InnerText);
        }

        private void ShowParams(XmlNode paramsNode)
        {
            _initializeFlag = true;

            dataGridViewMethodParameters.Rows.Clear();
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

        XmlNode GetNodeByName(List<XmlNode> list, string name)
        {
            foreach (XmlNode item in list)
            {
                if (item.Attributes["Caption"].InnerText == name)
                    return item;
            }
            throw (new ArgumentException("Node not found" + name));
        }
       
        private void dataGridViewMethods_SelectionChanged(object sender, EventArgs e)
        {
            if (true == _initializeFlag) return;
            if (dataGridViewMethods.SelectedCells.Count < 1) return;

            int rowIndex = dataGridViewMethods.SelectedCells[0].RowIndex;

            _initializeFlag = true;

            dataGridViewMethodParameters.Rows.Clear();

            DataGridViewRow selectRow = dataGridViewMethods.Rows[rowIndex];
            DataGridViewCell selectCell = dataGridViewMethods.Rows[rowIndex].Cells[0];

            if (null != selectCell.Tag)
            {
                XmlNode paramsNode = (XmlNode)selectCell.Tag;
                ShowParams(paramsNode);
            }

            _initializeFlag = false;
        }

        private void buttonResolve_Click(object sender, EventArgs e)
        {
            bool checkState = false;
            Button button = sender as Button;
            if (button.Text == "Check All")
                checkState = true;

            foreach (DataGridViewRow itemRow in dataGridViewMethods.Rows)
            {
                DataGridViewCheckBoxCell checkCell = itemRow.Cells[3] as DataGridViewCheckBoxCell;
                checkCell.Value = checkState;
            }
        }

        private void dataGridViewMethods_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            DataGridViewCheckBoxCell checkCell = dataGridViewMethods[e.ColumnIndex, e.RowIndex] as DataGridViewCheckBoxCell;
            if (null != checkCell)
                checkCell.Value = e.FormattedValue;
        }
        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow itemRow in dataGridViewMethods.Rows)
            {
                DataGridViewCheckBoxCell checkCell = itemRow.Cells[3] as DataGridViewCheckBoxCell;
                bool? value = checkCell.Value as bool?;
                if ((null != value) && value.Value == true)
                {
                    DataGridViewComboBoxCell comboCell = itemRow.Cells[4] as DataGridViewComboBoxCell;
                    List<XmlNode> list = comboCell.Tag as List<XmlNode>;
                    string name = comboCell.Value as string;
                    if( (null != name) && (name != "") )
                    {
                        if (true == IsScalar(name))
                        {
                            XmlNode paramsNode = itemRow.Cells[0].Tag as XmlNode;
                            XmlNode returnValueNode = paramsNode.SelectSingleNode("ReturnValue");
                            returnValueNode.Attributes["IsComProxy"].InnerText = "False";
                            returnValueNode.Attributes["IsExternal"].InnerText = "False";
                            returnValueNode.Attributes["Type"].InnerText = name;
                            returnValueNode.Attributes["TypeKind"].InnerText = "";
                            returnValueNode.Attributes["SourceKey"].InnerText = "";
                        }
                        else
                        {
                            XmlNode paramsNode = itemRow.Cells[0].Tag as XmlNode;
                            XmlNode returnValueNode = paramsNode.SelectSingleNode("ReturnValue");
                            XmlNode typeNode = GetNodeByName(list, name);

                            returnValueNode.Attributes["IsComProxy"].InnerText = "True";
                            returnValueNode.Attributes["IsExternal"].InnerText = "False";
                            returnValueNode.Attributes["Type"].InnerText = typeNode.Attributes["Caption"].InnerText;
                            returnValueNode.Attributes["TypeKind"].InnerText = "";
                            returnValueNode.Attributes["SourceKey"].InnerText = typeNode.ParentNode.ParentNode.Attributes["Key"].InnerText;
                        }
                  }
                }
            }
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            FillForm();
        }

    }
}
