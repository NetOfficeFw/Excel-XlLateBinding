using System;
using System.ComponentModel; 
using System.Collections.Generic;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.Visual
{
    internal enum RegKind
    {
        RegKind_Default = 0,
        RegKind_Register = 1,
        RegKind_None = 2
    }

    public partial class FormTypeLibBrowser : Form
    {
        #region Fields

        COMComponentReaderSettings _result;

        private string[] _selectedFiles;
        private List<TypeLibRegistryKey> _entriesList = new List<TypeLibRegistryKey>();
        private ListViewItemComparer     _typeLibSorter = new ListViewItemComparer(0);

        #endregion

        #region Properties

        public COMComponentReaderSettings Result
        {
            get 
            {
                return _result;
            }
        }
        
        #endregion

        #region Construction

        public FormTypeLibBrowser()
        {
            InitializeComponent();
            SetupListPropertySorter();
            ScanTypeLibRegistry();
            ShowResultItems();
        }

        #endregion

        #region Methods

        private void SetupListPropertySorter()
        {
            _typeLibSorter.AddColumnInfo(new ColumnTypeInfo(0, ColumnType.TypeInteger));
            _typeLibSorter.AddColumnInfo(new ColumnTypeInfo(1, ColumnType.TypeString));
            _typeLibSorter.AddColumnInfo(new ColumnTypeInfo(2, ColumnType.TypeString));
            _typeLibSorter.AddColumnInfo(new ColumnTypeInfo(3, ColumnType.TypeString));
            _typeLibSorter.AddColumnInfo(new ColumnTypeInfo(4, ColumnType.TypeString));

            listViewTypeLibInfo.ListViewItemSorter = _typeLibSorter;
        }

        private void ScanTypeLibRegistry()
        {
            _entriesList.Clear();

            foreach (TypeLibRegistryKey itemKey in TypeLibRegistry.Key.Keys)
            {
                _entriesList.Add(itemKey);
            }
        }

        private bool FilterIsMatched(bool filterEnabled, string name, string filterText)
        {
            if (filterEnabled == false) return true;
            int stringPosition = name.IndexOf(filterText, StringComparison.InvariantCultureIgnoreCase);
            if (stringPosition > -1)
                return true;
            else
                return false;
        }

        private void ShowResultItems()
        {
            string filterText = textBoxFilter.Text.Trim();
            bool filterEnabled = (filterText != "");

            int i = 1;
            listViewTypeLibInfo.Items.Clear();
            foreach (TypeLibRegistryKey itemKey in _entriesList)
            {
                foreach (TypeLibRegistryKey itemSubKey in itemKey.Keys)
                {
                    string version = itemSubKey.Name;
                    string name = "";
                    if(itemSubKey.Entries.Count > 0)
                       name = itemSubKey.Entries[0].Value.ToString();

                    if (true == FilterIsMatched(filterEnabled, name, filterText))
                    {
                        foreach (TypeLibRegistryKey itemSubSubKey in itemSubKey.Keys)
                        {
                            int iValue = -1;
                            int.TryParse(itemSubSubKey.Name, out iValue);
                            if (itemSubSubKey.Name == iValue.ToString())
                            {
                                foreach (TypeLibRegistryKey itemSubSubSubKey in itemSubSubKey.Keys)
                                {
                                    ListViewItem listItem = listViewTypeLibInfo.Items.Add(i.ToString());
                                    listItem.SubItems.Add(name);
                                    listItem.SubItems.Add(version);
                                    listItem.SubItems.Add(itemSubSubSubKey.Name);

                                    listItem.SubItems.Add(itemSubSubSubKey.Entries[0].Value.ToString());

                                    i++;
                                }
                            }

                        }
                    }

                }

            }

        }

        #endregion

        #region Gui Trigger

        private void FormTypelibBrowser_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                buttonCancel_Click(this, new EventArgs());
            }
        }

        private void buttonOkay_Click(object sender, EventArgs e)
        {
            if (listViewTypeLibInfo.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please choose one or more TypeLibs.", this.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _selectedFiles  = new string[listViewTypeLibInfo.SelectedItems.Count];
            int i=0;
            foreach (ListViewItem item in listViewTypeLibInfo.SelectedItems)
            {
                _selectedFiles[i] = item.SubItems[4].Text;
                i++;
            }

            labelProgress.Text = "Scan dependencies...";
            panelProgress.Visible = true;
            this.Refresh();
            XmlDocument dependDocument = Utils.DetectDependencies(_selectedFiles);
            panelProgress.Visible = false;

            FormDependencies formDepend = new FormDependencies(dependDocument);
            DialogResult dr = formDepend.ShowDialog(this);
            if (dr == DialogResult.OK)
            {
                int lenght1 = _selectedFiles.Length;
                int lenght2 = formDepend.ComponentsToGenerate.Length;
                int lenght3 = formDepend.ComponentsToNonGenerate.Length;
                
                int summaryLenght = lenght1 + lenght2 + lenght3;

                _result = new COMComponentReaderSettings();
                _result.CreateNewProject = !checkBoxAddToCurrentProject.Checked;
                _result.Files = new ComponentFile[summaryLenght];

                for (int l = 0; l < lenght1; l++)
                {
                    _result.Files[l] = new ComponentFile(_selectedFiles[l], false);
                }

                for (int l = lenght1; l < (lenght1 + lenght2); l++)
                {
                    _result.Files[l] = new ComponentFile(formDepend.ComponentsToGenerate[l - lenght1], false);
                }

                for (int l = lenght1 + lenght2; l < (lenght1 + lenght2 + lenght3); l++)
                {
                    _result.Files[l] = new ComponentFile(formDepend.ComponentsToNonGenerate[l - (lenght1 + lenght2)], true);
                }

                this.DialogResult = DialogResult.OK;
                this.Close();
            }

        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            _selectedFiles = new string[0];
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            ShowResultItems();
        }

        private void textBoxFilter_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                ShowResultItems();
        }

        private void listViewTypeLibInfo_Resize(object sender, EventArgs e)
        {
            listViewTypeLibInfo.Columns[0].Width = (this.Size.Width / 100) * 10;
            listViewTypeLibInfo.Columns[1].Width = (this.Size.Width / 100) * 30;
            listViewTypeLibInfo.Columns[2].Width = (this.Size.Width / 100) * 10;
            listViewTypeLibInfo.Columns[3].Width = (this.Size.Width / 100) * 10;
            listViewTypeLibInfo.Columns[4].Width = (this.Size.Width / 100) * 30;
        }

        private void listViewTypeLibInfo_Click(object sender, EventArgs e)
        {
            buttonOkay.Enabled = (listViewTypeLibInfo.SelectedItems.Count > 0);
        }

        private void listViewTypeLibInfo_DoubleClick(object sender, EventArgs e)
        {
            if (listViewTypeLibInfo.SelectedItems.Count == 0)
                return;
            buttonOkay_Click(this, new EventArgs());
        }

        private void listViewTypeLibInfo_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            _typeLibSorter.SortingColumn = e.Column;
            if (_typeLibSorter.SortingType == SortOrder.Ascending)
            {

                _typeLibSorter.SortingType = SortOrder.Descending;
                listViewTypeLibInfo.Sort();
            }
            else
            {
                _typeLibSorter.SortingType = SortOrder.Ascending;
                listViewTypeLibInfo.Sort();
            }
        }

        private void listViewTypeLibInfo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                buttonOkay_Click(this,new EventArgs());
            }
        }

        #endregion          
    }
}
