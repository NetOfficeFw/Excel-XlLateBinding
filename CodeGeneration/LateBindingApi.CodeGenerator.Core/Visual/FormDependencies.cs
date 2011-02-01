using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.Visual
{
    public partial class FormDependencies : Form
    {
        #region Construction

        public FormDependencies(XmlDocument dependDocument)
        {
            InitializeComponent();

            foreach (XmlNode nodeComponent in dependDocument.FirstChild.ChildNodes) 
            {
                TreeNode tn = treeViewComponents.Nodes.Add(nodeComponent.Attributes["Name"].InnerText);
                tn.Text += " " + nodeComponent.Attributes["File"].InnerText;

                foreach (XmlNode itemParentNode in nodeComponent.ChildNodes)
                {
                    TreeNode tnParent = tn.Nodes.Add("Used in: " + itemParentNode.Attributes["File"].InnerText);
                    tnParent.BackColor = Color.FromKnownColor(KnownColor.Control);                    
                }
            }
        }
        
        #endregion

        #region Properties

        internal  string[] ComponentsToGenerate
        {
            get 
            {
                List<string> returnList = new List<string>();
                foreach (TreeNode tnNode in treeViewComponents.Nodes )
                {
                    if (true == tnNode.Checked)
                    {
                        string addString = tnNode.Text.Substring(tnNode.Text.IndexOf(" ")).Trim();
                        returnList.Add(addString); 
                    }
                }
                return returnList.ToArray();
            }
        }

        internal string[] ComponentsToNonGenerate
        {
            get
            {
                List<string> returnList = new List<string>();
                foreach (TreeNode tnNode in treeViewComponents.Nodes)
                {
                    if (false == tnNode.Checked)
                    {
                        string addString = tnNode.Text.Substring(tnNode.Text.IndexOf(" ")).Trim();
                        returnList.Add(addString); 

                    }
                }
                return returnList.ToArray();
            }
        }
        
        #endregion

        #region Gui Trigger

        private void buttonOk_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        #endregion
    }
}
