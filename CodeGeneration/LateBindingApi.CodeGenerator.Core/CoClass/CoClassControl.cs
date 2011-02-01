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
    public partial class CoClassControl : UserControl
    {   
        #region Construction

        public CoClassControl()
        {
            InitializeComponent();
            this.Visible = false;
            detailsControlMain.Initialize();
            methodsControlMain.Initialize(false);
            propertiesControlMain.Initialize();
            eventsControlMain.Initialize();
        }
        
        #endregion

        #region Public Methods

        public void ShowItems(XmlNode classNode)
        {
            detailsControlMain.ShowItems(classNode);
            methodsControlMain.ShowItems(classNode);
            propertiesControlMain.ShowItems(classNode);
            eventsControlMain.ShowItems(classNode);
        }

        public void Clear()
        {
            detailsControlMain.Clear();
            methodsControlMain.Clear();
            propertiesControlMain.Clear();
            eventsControlMain.Clear();
        }

        #endregion
    }
}
