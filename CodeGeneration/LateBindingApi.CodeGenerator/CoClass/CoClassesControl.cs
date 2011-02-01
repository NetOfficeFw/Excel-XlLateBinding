using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator
{
    public partial class CoClassesControl : UserControl
    {
        #region Construction

        public CoClassesControl()
        {
            InitializeComponent();
            this.Visible = false;
        }

        #endregion

        #region Public Methods

        public void Clear()
        {
            labelClassesInfo.Text = "";
        }

        public void ShowItems(XmlNode classesNode)
        {
            XmlNode componentNode = classesNode.ParentNode.SelectSingleNode("Components");

            int countOfEnums = classesNode.ChildNodes.Count;
            int countOfComponents = componentNode.ChildNodes.Count;
            labelClassesInfo.Text = string.Format("{0} Classes in {1} Components.", countOfEnums, countOfComponents);
        }

        #endregion        
    }
}
