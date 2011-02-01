using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core
{
    public partial class CoClassesControl : UserControl
    {
        #region Fields

        bool _initializeFlag;

        #endregion

        #region Construction

        public CoClassesControl()
        {
            InitializeComponent();
            this.Visible = false;
        }
        #endregion

        #region Public Methods

        public void ShowItems(XmlNode itemsNode)
        {
            _initializeFlag = true; 
            int countOfEnums = itemsNode.ChildNodes.Count;
            labelClassesInfo.Text = string.Format("{0} Classes listed.", countOfEnums);
            _initializeFlag = false;
        }

        public void Clear()
        {
            _initializeFlag = true;
            labelClassesInfo.Text = "";
            _initializeFlag = false;
        }

        #endregion
    }
}
