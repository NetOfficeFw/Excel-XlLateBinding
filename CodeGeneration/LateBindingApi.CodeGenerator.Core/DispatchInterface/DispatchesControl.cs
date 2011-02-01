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
    public partial class DispatchesControl : UserControl
    {
        #region Fields

        bool _initializeFlag;

        #endregion

        #region Construction

        public DispatchesControl()
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
            labelInterfacesInfo.Text = string.Format("{0} Interfaces listed.", countOfEnums);
            _initializeFlag = false;
        }

        public void Clear()
        {
            _initializeFlag = true;
            labelInterfacesInfo.Text = "";
            _initializeFlag = false;
        }

        #endregion
    }
}
