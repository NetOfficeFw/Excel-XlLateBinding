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
    public partial class DispatchesControl : UserControl
    {
        public DispatchesControl()
        {
            InitializeComponent();
            this.Visible = false;
        }

        public void ShowItems(XmlNode interfacesNode) 
        {
            XmlNode componentNode = interfacesNode.ParentNode.SelectSingleNode("Components");

            int countOfInterfaces = interfacesNode.ChildNodes.Count;
            int countOfComponents = componentNode.ChildNodes.Count;
            labelInterfacesInfo.Text = string.Format("{0} Interfaces in {1} Components.", countOfInterfaces, countOfComponents);
 
        }

        public void Clear()
        {
            labelInterfacesInfo.Text = "";
        }
    }
}
