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
    public partial class EventsControl : UserControl
    {
        #region Fields

        bool _calledFromFilter;
        bool _initializeFlag;
        // XmlNode _methodsNode;
        List<XmlNode> _nodesList = new List<XmlNode>();
        #endregion

        #region Construction

        public EventsControl()
        {
            InitializeComponent();
            Initialize();
        }

        #endregion

        #region Public Methods

        public void Clear()
        {
            dataGridViewEvents.Rows.Clear();
            dataGridViewEventsParams.Rows.Clear();
        }

        public void Initialize()
        {
            SetupGridColumns();
        }

        public void ShowItems(XmlNode methodsNode, bool clearOldResult)
        {
            _initializeFlag = true;

            if (true == clearOldResult)
            {
                _nodesList.Clear();
                dataGridViewEvents.Rows.Clear();
                dataGridViewEventsParams.Rows.Clear();
            }
    
            if(_calledFromFilter == false)
                _nodesList.Add(methodsNode);

            foreach (XmlNode methodNode in methodsNode.ChildNodes)
            {
                if (methodNode.Attributes["Name"].InnerText.IndexOf(textBoxMethodFilter.Text, StringComparison.InvariantCultureIgnoreCase) > -1)
                {
                    dataGridViewEvents.Rows.Add();
                    DataGridViewRow newRow = dataGridViewEvents.Rows[dataGridViewEvents.Rows.Count - 1];
                    newRow.Cells[0].Value = methodNode.Attributes["Name"].InnerText;
                    newRow.Cells[1].Value = methodNode.Attributes["DispId"].InnerText;
                    newRow.Cells[2].Value = GetComponentVersionAttributes(methodNode.SelectSingleNode("Components"));
                    newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                    newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
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
            dataGridViewEvents.Columns.Clear();
            dataGridViewEvents.Columns.Add("Name", "Name");
            dataGridViewEvents.Columns.Add("DispId", "DispId");
            dataGridViewEvents.Columns.Add("Components", "Components");
            dataGridViewEvents.Columns[0].ReadOnly = true;
            dataGridViewEvents.Columns[1].ReadOnly = true;
            dataGridViewEvents.Columns[2].ReadOnly = true;

            dataGridViewEventsParams.Columns.Clear();
            dataGridViewEventsParams.Columns.Add("Type", "Type");
            dataGridViewEventsParams.Columns.Add("Name", "Name");
            dataGridViewEventsParams.Columns.Add("Optional", "Optional");

            dataGridViewEventsParams.Columns[0].ReadOnly = true;
            dataGridViewEventsParams.Columns[1].ReadOnly = true;
            dataGridViewEventsParams.Columns[2].ReadOnly = true;
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
            if (dataGridViewEvents.SelectedCells.Count < 1) return;

            _initializeFlag = true;

            dataGridViewEventsParams.Rows.Clear();
            DataGridViewCell selectCell = dataGridViewEvents.SelectedCells[0];
            DataGridViewRow selectRow = dataGridViewEvents.Rows[selectCell.RowIndex];
            XmlNode methodNode = (XmlNode)selectRow.Tag;
            XmlNode paramsNode = methodNode.SelectSingleNode("Parameters");
            foreach (XmlNode paramNode in paramsNode.ChildNodes)
            {
                dataGridViewEventsParams.Rows.Add();
                DataGridViewRow newRow = dataGridViewEventsParams.Rows[dataGridViewEventsParams.Rows.Count - 1];
                newRow.Cells[0].Value = paramNode.Attributes["Type"].InnerText;
                newRow.Cells[1].Value = paramNode.Attributes["Name"].InnerText;
                newRow.Cells[2].Value = paramNode.Attributes["Optional"].InnerText;
                newRow.Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
                newRow.Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.Control);

                newRow.Tag = paramNode;
            }

            _initializeFlag = false;
        }

        private void textBoxMethodFilter_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Return)
            {
                _calledFromFilter = true;
                foreach (XmlNode itemNode in _nodesList)
                {
                    ShowItems(itemNode, false);
                }
                _calledFromFilter = false;
            }
        }

        #endregion
    }
}
