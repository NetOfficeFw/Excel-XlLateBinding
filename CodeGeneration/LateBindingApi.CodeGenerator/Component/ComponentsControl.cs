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
    public partial class ComponentsControl : UserControl
    {
        #region Fields

        COMComponentReader _reader;
        bool               _initializeFlag;

        #endregion

        #region Construction

        public ComponentsControl()
        {
            InitializeComponent();
            this.Visible = false;
            SetupGridColumns();
        }
        
        #endregion

        #region Public Methods

        public void Clear()
        {
            textBoxSolutionName.Clear(); 
            textBoxClassPrefix.Clear();
            dataGridViewComponents.Rows.Clear();
        }

        public void ShowItems(COMComponentReader reader)
        {
            _initializeFlag = true;
            _reader = reader;
           dataGridViewComponents.Rows.Clear();
           foreach (COMComponent itemComponent in reader.Components)
           {
               dataGridViewComponents.Rows.Add();
               int rowIndex = dataGridViewComponents.Rows.Count - 1;

               dataGridViewComponents.Rows[rowIndex].Tag = itemComponent;

               dataGridViewComponents.Rows[rowIndex].Cells[0].Value = itemComponent.Name;
               dataGridViewComponents.Rows[rowIndex].Cells[0].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[0].ReadOnly = true;
              
               dataGridViewComponents.Rows[rowIndex].Cells[1].Value = itemComponent.Description;
               dataGridViewComponents.Rows[rowIndex].Cells[1].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki);
               
               dataGridViewComponents.Rows[rowIndex].Cells[2].Value = itemComponent.VersionAttribute;
               dataGridViewComponents.Rows[rowIndex].Cells[2].Style.BackColor = Color.FromKnownColor(KnownColor.DarkKhaki);

               dataGridViewComponents.Rows[rowIndex].Cells[3].Value = itemComponent.ContainingFile;
               dataGridViewComponents.Rows[rowIndex].Cells[3].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[3].ReadOnly = true;

               dataGridViewComponents.Rows[rowIndex].Cells[4].Value = itemComponent.GUID;
               dataGridViewComponents.Rows[rowIndex].Cells[4].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[4].ReadOnly = true;

               dataGridViewComponents.Rows[rowIndex].Cells[5].Value = itemComponent.MajorVersion;
               dataGridViewComponents.Rows[rowIndex].Cells[5].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[5].ReadOnly = true;
 
               dataGridViewComponents.Rows[rowIndex].Cells[6].Value = itemComponent.MinorVersion;
               dataGridViewComponents.Rows[rowIndex].Cells[6].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[6].ReadOnly = true;

               dataGridViewComponents.Rows[rowIndex].Cells[7].Value = itemComponent.LCID;
               dataGridViewComponents.Rows[rowIndex].Cells[7].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[7].ReadOnly = true;

               dataGridViewComponents.Rows[rowIndex].Cells[8].Value = itemComponent.SysKind;
               dataGridViewComponents.Rows[rowIndex].Cells[8].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[8].ReadOnly = true;

               dataGridViewComponents.Rows[rowIndex].Cells[9].Value = itemComponent.ContainingFile;
               dataGridViewComponents.Rows[rowIndex].Cells[9].Style.BackColor = Color.FromKnownColor(KnownColor.Control);
               dataGridViewComponents.Rows[rowIndex].Cells[9].ReadOnly = true;
           }

           XmlNode solutionNode = reader.COMTree.SelectSingleNode("LateBindingApi.CodeGenerator.Document/Solution");
           if (null == solutionNode)
               solutionNode = CreateSolutionNode();

           textBoxSolutionName.Text = solutionNode.Attributes["Name"].Value;
           textBoxClassPrefix.Text = solutionNode.Attributes["Prefix"].Value;

           _initializeFlag = false;
        }
        
        #endregion

        #region Private Methods

        private COMComponent GetComponentByKey(string key)
        {
            foreach (COMComponent itemComponent in _reader.Components)
            {
                if (itemComponent.Key == key)
                    return itemComponent;
            }

            throw (new ArgumentOutOfRangeException("Key not found. " + key));
        }

        private XmlNode CreateSolutionNode()
        {

            XmlNode solutionNode = _reader.COMTree.CreateElement("Solution");
            XmlAttribute attrib = _reader.COMTree.CreateAttribute("Name");
            attrib.InnerText = "LateBindingApi." + _reader.Components[0].Name;
            solutionNode.Attributes.Append(attrib);

            attrib = _reader.COMTree.CreateAttribute("Prefix");
            attrib.InnerText = _reader.Components[0].Name;
            solutionNode.Attributes.Append(attrib);

            _reader.COMTree.SelectSingleNode("LateBindingApi.CodeGenerator.Document").AppendChild(solutionNode);


            return solutionNode;
        }

        private void SetupGridColumns()
        {
            dataGridViewComponents.Columns.Clear();
 
            dataGridViewComponents.Columns.Add("Name", "Name");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count -1].Width = 55;

            dataGridViewComponents.Columns.Add("Description", "Description");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].Width = 155;

            dataGridViewComponents.Columns.Add("Version", "Version");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].Width = 55;

            dataGridViewComponents.Columns.Add("ContainingFile", "ContainingFile");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;             
            
            dataGridViewComponents.Columns.Add("GUID", "GUID");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].Width = 55;                       
            
            dataGridViewComponents.Columns.Add("Major", "Major");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].Width = 55;            
            
            dataGridViewComponents.Columns.Add("Minor", "Minor");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].Width = 55;          
            
            dataGridViewComponents.Columns.Add("LCID", "LCID");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].Width = 55;
            
            dataGridViewComponents.Columns.Add("SysKind", "SysKind");
            dataGridViewComponents.Columns[dataGridViewComponents.Columns.Count - 1].Width = 100;

            dataGridViewComponents.Columns.Add("HelpFile", "HelpFile");  
        }
        
        #endregion

        #region Gui Trigger

        private void buttonCancel_Click(object sender, EventArgs e)
        {   
            try
            {
                SetupGridColumns();
                ShowItems(_reader);
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }

        private void buttonApply_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (DataGridViewRow itemRow in dataGridViewComponents.Rows)
                {
                    string key = ((COMComponent)itemRow.Tag).Key;
                    COMComponent component = GetComponentByKey(key);
                    component.Description = (string)itemRow.Cells[1].Value;
                    component.VersionAttribute = (string) itemRow.Cells[2].Value;
                }

            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }

        private void textBoxSolutionName_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (true == _initializeFlag) return;

                XmlNode solutionNode = _reader.COMTree.SelectSingleNode(Constants.Solution);
                solutionNode.Attributes["Name"].Value = textBoxSolutionName.Text.Trim();
                solutionNode.Attributes["Prefix"].Value = textBoxClassPrefix.Text.Trim();

            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }

        private void dataGridViewComponents_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if((e.RowIndex < 0) || (true==_initializeFlag)) return;

                DataGridViewRow selectRow = dataGridViewComponents.Rows[e.RowIndex];
                COMComponent selectComponent = (COMComponent)selectRow.Tag;

                selectComponent.Description       = (string) selectRow.Cells[1].Value;
                selectComponent.VersionAttribute  = (string)selectRow.Cells[2].Value;
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }

        #endregion
    }
}
