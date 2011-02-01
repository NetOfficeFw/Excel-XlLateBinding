using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using LateBindingApi.Excel;
using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Addins;
using LateBindingApi.Excel.COMAddins;
using LateBindingApi.Excel.XlRegistry;

namespace TestApp
{

    public partial class Form1 : Form
    {
        XlApplication excelApplication;

        #region Construction

        public Form1()
        {
            InitializeComponent();
            excelApplication = new XlApplication();
            excelApplication.Workbooks.Add();
            excelApplication.Visible = true;
            RefreshView(); 
        }

        #endregion

        private void RefreshView()
        {
            ScanExcelObjectModel();
            SetupListViewCOMAddins();
            SetupInfo();
            SetupSettings();
            SetupMachineRegistry();
            SetupUserRegistry();
            SetupListViewAddins();
            treeViewApplicationTree.Nodes[0].Expand();
          
        }

        private object CallProperty(object targetObject, string propertyName)
        {
            try
            {
                Type objectType = targetObject.GetType();
                System.Reflection.PropertyInfo propInfo = objectType.GetProperty(propertyName);
                object propValue = propInfo.GetValue(targetObject, null);
                return propValue;
            }
            catch
            { 
                return null;
            }
        }

        private string[] GetAllIXlPropertyNames(object frameworkObject)
        {
            if (!(frameworkObject is IXlObject))
                throw (new ArgumentException("Parameter is not IXlObject", "frameworkObject"));

            List<string> nameList = new List<string>();

            Type objectType = frameworkObject.GetType();

            foreach (System.Reflection.PropertyInfo proInfo in objectType.GetProperties())
            {

                Type baseInterface = proInfo.PropertyType.GetInterface("IXlObject");
                if(null!=baseInterface)
                    nameList.Add(proInfo.Name);
            }

            return nameList.ToArray();
        }

        private void ScanExcelObjectModel()
        {
            treeViewApplicationTree.Nodes.Clear();
            propertyGridDetails.SelectedObject = excelApplication;
            TreeNode tnRoot = treeViewApplicationTree.Nodes.Add(excelApplication.ToString());
            tnRoot.Tag = excelApplication;
            string[] propertyNames = GetAllIXlPropertyNames(excelApplication);
            foreach (string propName in propertyNames)
            {
                TreeNode tnProperty = tnRoot.Nodes.Add(propName);
                tnProperty.Nodes.Add("Please Wait...");
            }

        }

        private void SetupApplicationGrid()
        {
            propertyGridDetails.SelectedObject = excelApplication;
        }
  
        private void SetupListViewCOMAddins()
        {
            listViewCOMAddins.Items.Clear();
            foreach (XlCOMAddin comAddin in excelApplication.COMAddIns)
            {
                ListViewItem lviAddin = listViewCOMAddins.Items.Add(comAddin.ProgId);
                lviAddin.SubItems.Add(comAddin.Connect.ToString());
                lviAddin.SubItems.Add(comAddin.Description);
                lviAddin.SubItems.Add(comAddin.Guid);
            }

        }

        private void SetupListViewAddins()
        {
            listViewAddins.Items.Clear();
            foreach (XlAddin addin in excelApplication.AddIns)
            {

                ListViewItem lviAddin = listViewAddins.Items.Add(addin.Name);
                lviAddin.SubItems.Add(addin.FullName);
                lviAddin.SubItems.Add(addin.Path);
                lviAddin.SubItems.Add(addin.ProgId);
                lviAddin.SubItems.Add(addin.Installed.ToString());
                lviAddin.SubItems.Add(addin.CLSID);
            }

        }

        private void SetupInfo()
        {
             
            ListViewItem lviAddin = listViewInfo.Items.Add("FrameworkVersion");
            lviAddin.SubItems.Add(XlLateBindingApiSettings.FrameworkVersion.ToString());

            lviAddin = listViewInfo.Items.Add("DefaultThreadLCID");
            lviAddin.SubItems.Add(XlLateBindingApiSettings.DefaultThreadLCID);
             
        }

        private void SetupSettings()
        {
            ListViewItem lviAddin = listViewSettings.Items.Add("XlThreadCulture");
            lviAddin.SubItems.Add(LateBindingApi.Excel.XlLateBindingApiSettings.XlThreadCulture.ToString());

        }

        private void SetupUserRegistry()
        {
            
            treeViewRegistryUser.Nodes.Clear();
            listViewRegistryUser.Items.Clear();

            if (true == XlRegistryCurrentUser.Exists)
            {

               
                TreeNode tnRootNode = treeViewRegistryUser.Nodes.Add(XlRegistryCurrentUser.Key.Name);
                tnRootNode.Tag = XlRegistryCurrentUser.Key.Entries;
                ApplyKey(tnRootNode, XlRegistryCurrentUser.Key.Keys);

            }

            if(treeViewRegistryUser.Nodes.Count > 0)
                treeViewRegistryUser.Nodes[0].Expand();


        }

        private void SetupMachineRegistry()
        {
            treeViewRegistryMachine.Nodes.Clear();
            listViewRegistryMachine.Items.Clear();

            if (true == XlRegistryLocalMachine.Exists)
            {
                TreeNode tnRootNode = treeViewRegistryMachine.Nodes.Add(XlRegistryLocalMachine.Key.Name);
                tnRootNode.Tag = XlRegistryLocalMachine.Key.Entries;
                ApplyKey(tnRootNode, XlRegistryLocalMachine.Key.Keys);

            }

            if (treeViewRegistryMachine.Nodes.Count > 0)
                treeViewRegistryMachine.Nodes[0].Expand();

        }

        private void ApplyKey(TreeNode tnRootNode, XlRegistryKeys regKeys)
        {
            foreach (XlRegistryKey regKey in regKeys)
            {
                TreeNode tnSubNode = tnRootNode.Nodes.Add(regKey.Name);
                tnSubNode.Tag = regKey.Entries;
                ApplyKey(tnSubNode, regKey.Keys);
            }
        }

        #region Gui Trigger

        private void treeViewApplicationTree_AfterExpand(object sender, TreeViewEventArgs e)
        {
            if ((null != e.Node) && (null == e.Node.Tag))
            {
                TreeNode parentNode = e.Node.Parent;
                if (parentNode.Tag != null)
                {
                    e.Node.Nodes.Clear();
                    object targetProperty = CallProperty(parentNode.Tag, e.Node.Text);
                    if (targetProperty != null)
                    {
                        propertyGridDetails.SelectedObject = targetProperty;
                        string[] propertyNames = GetAllIXlPropertyNames(targetProperty);
                        foreach (string propName in propertyNames)
                        {
                            TreeNode tnProperty = e.Node.Nodes.Add(propName);
                        }
                    }
                }
                
            }
        }

        private void treeViewApplicationTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if ((null != e.Node) && (null != e.Node.Tag))
            {
                propertyGridDetails.SelectedObject = e.Node.Tag;
            }
            else
            {
                TreeNode parentNode = e.Node.Parent;
                if (parentNode.Tag != null)
                {
                    object targetProperty = CallProperty(parentNode.Tag, e.Node.Text);
                    if(targetProperty!=null)
                    {
                        propertyGridDetails.SelectedObject = targetProperty;
                    }
                    else
                        propertyGridDetails.SelectedObject = null;
                }
                else
                {
                    propertyGridDetails.SelectedObject = null;
                }
            }
        }
   
        private void treeViewRegistryUser_AfterSelect(object sender, TreeViewEventArgs e)
        {

            if (treeViewRegistryUser.SelectedNode != null)
            {
                TreeNode tnNode = treeViewRegistryUser.SelectedNode;

                listViewRegistryUser.Items.Clear();
                XlRegistryEntries entries = (XlRegistryEntries)tnNode.Tag;
                if (null != entries)
                { 
                    foreach (XlRegistryEntry entry in entries)
                    {
                        ListViewItem lvi = listViewRegistryUser.Items.Add(entry.Name);
                        lvi.SubItems.Add(entry.Value.ToString());
                        lvi.SubItems.Add(entry.ValueType.ToString());
                    }
                }
            }
           

        }

        private void treeViewRegistryMachine_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if (treeViewRegistryMachine.SelectedNode != null)
            {
                TreeNode tnNode = treeViewRegistryMachine.SelectedNode;

                listViewRegistryMachine.Items.Clear();
                XlRegistryEntries entries = (XlRegistryEntries)tnNode.Tag;
                if (null != entries)
                {
                    foreach (XlRegistryEntry entry in entries)
                    {
                        ListViewItem lvi = listViewRegistryMachine.Items.Add(entry.Name);
                        lvi.SubItems.Add(entry.Value.ToString());
                        lvi.SubItems.Add(entry.ValueType.ToString());
                    }
                }
            }
        }

        private void toolStripButtonCloseApplication_Click(object sender, EventArgs e)
        {
            excelApplication.Quit();
            excelApplication.Dispose();
            this.Close();
        }
    
        private void itemAddNewWorkbook_Click(object sender, EventArgs e)
        {
            excelApplication.Workbooks.Add();
        }

        private void itemCloseApplication_Click(object sender, EventArgs e)
        {
            excelApplication.Quit();
            excelApplication.Dispose();
            this.Close();
        }

        private void itemCloseWorkbook_Click(object sender, EventArgs e)
        {
            if (excelApplication.ActiveWorkbook != null)
                excelApplication.ActiveWorkbook.Close(false);
        }

        private void toolStripButtonRefresh_Click(object sender, EventArgs e)
        {
            RefreshView();
        }

        private void itemOpenWorkbook_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = Environment.CurrentDirectory;
            ofd.Filter = "Bla(*.*)|*.*";
            DialogResult dr = ofd.ShowDialog(this);
            if (dr != DialogResult.Cancel)
            {
                excelApplication.Workbooks.Open(ofd.FileName);
            }
        }

        #endregion
    }
}
