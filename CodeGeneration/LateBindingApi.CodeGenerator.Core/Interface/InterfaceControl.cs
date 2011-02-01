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
    public partial class InterfaceControl : UserControl
    {    
        #region Construction

        public InterfaceControl()
        {
            InitializeComponent();
            this.Visible = false;
            detailsControlMain.Initialize();
            methodsControlMain.Initialize(true);
            propertiesControlMain.Initialize();
        }
        
        #endregion
        
        #region Public Methods

        public void ShowItems(XmlNode interfaceNode)
        {
            detailsControlMain.ShowItems(interfaceNode);
            methodsControlMain.ShowItems(interfaceNode);
            propertiesControlMain.ShowItems(interfaceNode);
        }

        public void Clear()
        {
            detailsControlMain.Clear();
            methodsControlMain.Clear();
            propertiesControlMain.Clear();
        }
        
        #endregion
    }
}
