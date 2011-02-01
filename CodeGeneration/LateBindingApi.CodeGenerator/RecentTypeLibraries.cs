using System;
using System.Windows.Forms;
using System.Collections;
using System.Xml;
using System.Text;
using LateBindingApi.CodeGenerator.Core;

namespace LateBindingApi.CodeGenerator
{
    public class RecentTypeLibraries
    {
        #region Member

        private readonly string _fileName = "RecentTypeLibraries.xml";
        XmlDocument _document = new XmlDocument();
        EventHandler _recentClickHandler;
        ToolStripMenuItem _toolstrip;
        
        #endregion

        #region Construction
        
        public RecentTypeLibraries(EventHandler recentClickHandler, ToolStripMenuItem toolstrip)
        {
            _recentClickHandler = recentClickHandler;
            _toolstrip = toolstrip;
        }

        #endregion

        #region Methods

        public void Add(XmlNode componentsNode)
        {
            foreach (XmlNode componentNode in componentsNode.ChildNodes)
            {
                string text = componentNode.Attributes["Description"].InnerText;
                string fileName = componentNode.Attributes["ContainingFile"].InnerText;

                foreach (ToolStripItem item in _toolstrip.DropDownItems)
                {
                    if (item.Text == text)
                        return;
                }

                if (_toolstrip.DropDownItems.Count < 7)
                {
                    ToolStripItem item = _toolstrip.DropDownItems.Add(text);
                    _toolstrip.DropDownItems.Insert(2, item);
                    item.Tag = fileName;
                    item.Click += _recentClickHandler;

                    XmlNode newNode = _document.CreateElement(System.Xml.XmlConvert.EncodeName(text));
                    newNode.InnerText = fileName;
                    _document.FirstChild.AppendChild(newNode);
                }
                else
                {
                    string encodedNodeName = XmlConvert.EncodeName(_toolstrip.DropDownItems[5].Text);
                    XmlNode deleteNode = _document.SelectSingleNode("/TypeLibraries/" + encodedNodeName);
                    _document.FirstChild.RemoveChild(deleteNode);

                    _toolstrip.DropDownItems.Remove(_toolstrip.DropDownItems[6]);
                    ToolStripItem item = _toolstrip.DropDownItems.Add(text);
                    _toolstrip.DropDownItems.Insert(2, item);
                    item.Text = text;
                    item.Tag = fileName;

                    XmlNode newNode = _document.CreateElement(System.Xml.XmlConvert.EncodeName(text));
                    newNode.InnerText = fileName;
                    _document.FirstChild.AppendChild(newNode);
                }
            }
        }

        public void LoadFromFile()
        {
            string fullFileName = System.IO.Path.Combine(Environment.CurrentDirectory, _fileName);
            if (true == System.IO.File.Exists(fullFileName))
            {
                _document.Load(fullFileName);
            }
            else
            {
                XmlNode newNode = _document.CreateElement("TypeLibraries");
                _document.AppendChild(newNode);
            }

            foreach (XmlNode itemNode in _document.FirstChild.ChildNodes)
            {
                ToolStripItem item = _toolstrip.DropDownItems.Add(System.Xml.XmlConvert.DecodeName(itemNode.Name));
                _toolstrip.DropDownItems.Insert(2, item);

                item.Tag = itemNode.InnerText;
                item.Click += _recentClickHandler; 
            }
        }

        public void SaveToFile()
        {
            string fullFileName = System.IO.Path.Combine(Environment.CurrentDirectory, _fileName);
            if (true == System.IO.File.Exists(fullFileName))
                System.IO.File.Delete(fullFileName);

            _document.Save(fullFileName);
        }

        #endregion
    }
}
