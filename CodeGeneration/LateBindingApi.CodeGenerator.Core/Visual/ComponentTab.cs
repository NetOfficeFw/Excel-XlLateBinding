using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.Visual
{
    public partial class ComponentTab : UserControl
    {
        public ComponentTab()
        {
            InitializeComponent();
        }

        public void Clear()
        {
            itemEntitiesGridMain.Clear();
            methodsGridMain.Clear();
            propertiesGridMain.Clear();
            xmlControlMain.Clear();
            eventsGridMain.Clear();
        }

        public void Show(ItemFilter filter, XmlNode node)
        {
            this.Visible = true;
            itemEntitiesGridMain.Show(node);
            methodsGridMain.Show(filter, node);
            propertiesGridMain.Show(filter, node);
            xmlControlMain.Show(node);
            eventsGridMain.Show(filter, node);
        }
    }
}
