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
    public partial class ItemEntitiesGrid : UserControl
    {
        public ItemEntitiesGrid()
        {
            InitializeComponent();
        }

        public void Clear()
        {
            attributesGrid.Clear();
            refComponentsGrid.Clear();
            refInterfacesGrid.Clear();
        }

        public void Show(XmlNode node)
        {
            attributesGrid.Show(node);
            refComponentsGrid.Show(node);
            refInterfacesGrid.Show(node, "InheritedInterfaces/RefInterface");
        }

    }
}
