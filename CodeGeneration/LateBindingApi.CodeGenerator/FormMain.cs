using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using LateBindingApi.CodeGenerator.Core;
using LateBindingApi.CodeGenerator.Core.Visual;
 
namespace LateBindingApi.CodeGenerator.WFApplication
{
    public partial class FormMain : Form
    {
        #region Fields

        ItemFilter         _filter    = new ItemFilter();
        COMComponentReader _comReader = new COMComponentReader();

        #endregion

        #region Construction

        public FormMain()
        {
            InitializeComponent();
            this.Text = this.GetType().Assembly.GetName().Name;

            foreach (Control itemControl in splitContainerMain.Panel2.Controls)
                itemControl.Dock = DockStyle.Fill;            

            componentTreeViewMain.AfterSelect += new TreeViewEventHandler(componentTreeViewMain_AfterSelect);  
        }

        void componentTreeViewMain_AfterSelect(object sender, TreeViewEventArgs e)
        {
            foreach (Control itemControl in splitContainerMain.Panel2.Controls)
                itemControl.Visible = false;

            XmlNode node = e.Node.Tag as XmlNode;
            switch (node.Name)
            { 
                case "Components":
                    childAttributesGridMain.Show(node);
                    break;
                case "Component":
                    attributesGridMain.Show(node);
                    break;
                case "Externals":
                    childAttributesGridMain.Show(node);
                    break;
                case "External":
                    attributesGridMain.Show(node);
                    break;
                case "ExternalComponents":
                    childAttributesGridMain.Show(node);
                    break;
                case "ExternalComponent":
                    attributesGridMain.Show(node);
                    break;
                case "Solution":
                    attributesGridMain.Show(node);
                    break;
                case "Project":
                    attributesGridMain.Show(node);
                    break;
                case "Enums":
                    
                    break;
                case "Enum":
                    enumGridMain.Show(node);
                    break;

                case "CoClasses":
                    
                    break;
                case "CoClass":
                    componentTabMain.Show(_filter, node);
                    break;

                case "DispatchInterfaces":

                    break;
                case "DispatchInterface":
                    componentTabMain.Show(_filter, node);
                    break;

                case "Interfaces":

                    break;
                case "Interface":
                    componentTabMain.Show(_filter, node);
                    break;
            }
            
           

        }

        #endregion

        #region Menu Click Trigger

        private void menuItemLoadTypeLibrary_Click(object sender, EventArgs e)
        {
            FormTypeLibBrowser browseForm = new FormTypeLibBrowser();
            browseForm.ShowDialog(this);
            if (DialogResult.OK == browseForm.DialogResult)
            { 
                _comReader.LoadCOMComponents(browseForm.Result);
                OnTypeLibraryLoaded();
            }
        }

        private void menuItemLoadProject_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog fileDialog = new OpenFileDialog();
                fileDialog.Filter = "ProjectFiles|*.xml";
                DialogResult dr = fileDialog.ShowDialog(this);
                if (dr == DialogResult.OK)
                {
                    _comReader.LoadProject(fileDialog.FileName);
                    OnTypeLibraryLoaded();
                }
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException, "Failed to load Project Information.");
                errorForm.ShowDialog(this);
            }
        }

        private void menuItemSaveProject_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog fileDialog = new SaveFileDialog();
                fileDialog.Filter = "ProjectFiles|*.xml";
                DialogResult dr = fileDialog.ShowDialog(this);
                if (dr == DialogResult.OK)
                {
                    _comReader.SaveProject(fileDialog.FileName);
                    MessageBox.Show("Project successfully saved.", this.GetType().Assembly.GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException, "Failed to save Project Information.");
                errorForm.ShowDialog(this);
            }
        }
        
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception throwedException)
            {
                FormShowError errorForm = new FormShowError(throwedException);
                errorForm.ShowDialog(this);
            }
        }

        private void menuItemGenerateCore_Click(object sender, EventArgs e)
        {

        }

        private void menuItemGenerateCode_Click(object sender, EventArgs e)
        {

        }

        private void websiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormAbout formAbout = new FormAbout();
            formAbout.ShowDialog(this);
        }

        private void toolStripFilterAll_CheckedChanged(object sender, EventArgs e)
        {
            if (toolStripFilterAll.Checked == true)
            {
                toolStripFilterConflicts.Checked = false;
                toolStripFilterNoConflicts.Checked = false;
                _filter.Mode = ItemFilter.FilterMode.All;
                componentTreeViewMain.SaveCurrentNodePath(); 
                OnTypeLibraryLoaded();
                componentTreeViewMain.RestoreExpandState();
            }
        }

        private void toolStripFilterConflicts_CheckedChanged(object sender, EventArgs e)
        {
            if (toolStripFilterConflicts.Checked == true)
            {
                toolStripFilterAll.Checked = false;
                toolStripFilterNoConflicts.Checked = false;
                _filter.Mode = ItemFilter.FilterMode.Conflict;
                componentTreeViewMain.SaveCurrentNodePath(); 
                OnTypeLibraryLoaded();
                componentTreeViewMain.RestoreExpandState();
            }
        }

        private void toolStripFilterNoConflicts_CheckedChanged(object sender, EventArgs e)
        {
            if (toolStripFilterNoConflicts.Checked == true)
            {
                toolStripFilterAll.Checked = false;
                toolStripFilterConflicts.Checked = false;
                _filter.Mode = ItemFilter.FilterMode.NoConflict; 
                componentTreeViewMain.SaveCurrentNodePath(); 
                OnTypeLibraryLoaded();
                componentTreeViewMain.RestoreExpandState();
            }
        }

        #endregion

        private void OnTypeLibraryLoaded()
        {
            componentTreeViewMain.Show(_comReader.Document.DocumentElement); 

            int countOfLibraries = _comReader.Document.SelectNodes(XPathConstants.Components + "/Component").Count;
            if (countOfLibraries > 0)
            {
                this.Text = string.Format("{0} - {1} Libraries loaded", this.GetType().Assembly.GetName().Name, countOfLibraries.ToString() );
                componentTreeViewMain.SelectFirstNode();
                menuItemSaveProject.Enabled = true;            
                menuItemGenerateCode.Enabled = true;
                splitContainerMain.Visible = true;            
            }
            else
            {
                this.Text = this.GetType().Assembly.GetName().Name;
                menuItemSaveProject.Enabled = false;
                menuItemGenerateCode.Enabled = false;
                splitContainerMain.Visible = false;                
            }
        }
   }
}