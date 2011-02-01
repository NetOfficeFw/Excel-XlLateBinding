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
    public partial class EnumsControl : UserControl
    {
        #region Construction

        public EnumsControl()
        {
            InitializeComponent();
            this.Visible = false;
        }
        
        #endregion

        #region Public Methods

        public void Clear()
        {
        }

        public void ShowItems(XmlNode enums)
        {
           XmlNode componentNode = enums.ParentNode.SelectSingleNode("Components");

            int countOfEnums = enums.ChildNodes.Count;
            int countOfComponents = componentNode.ChildNodes.Count;
            labelEnumsInfo.Text = string.Format("{0} Enums in {1} Components.", countOfEnums, countOfComponents);
        }
        
        #endregion
    }
}
