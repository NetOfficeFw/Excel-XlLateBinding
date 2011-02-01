using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

using LateBindingApi.CodeGenerator.Core;

namespace LateBindingApi.CodeGenerator.Core.Visual
{
    public partial class XmlControl : UserControl
    {
        XmlNode _node;

        public XmlControl()
        {
            InitializeComponent();
        }

        public void Clear()
        {
            textBoxContent.Clear();
        }

        public void Show(XmlNode node)
        {
            _node = node;
            textBoxSearch.Text = ""; 
            textBoxContent.Text = Utils.PrintXML(_node); 
        }

        private void buttonReset_Click(object sender, EventArgs e)
        {
            textBoxContent.Text = Utils.PrintXML(_node); 
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            _node.InnerText = textBoxContent.Text;
        }

        private void textBoxSearch_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Return)
            {

            }
        }

    }
}
