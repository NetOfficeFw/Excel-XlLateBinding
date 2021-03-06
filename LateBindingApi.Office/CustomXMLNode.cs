//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class CustomXMLNode : _IMsoDispObj
	{
		#region Construction

		public CustomXMLNode(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public CustomXMLNode(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public CustomXMLNode(COMObject replacedObject) : base(replacedObject)
		{
		}

		public CustomXMLNode()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNodes Attributes
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Attributes");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLNodes newClass = new LateBindingApi.Office.CustomXMLNodes(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string BaseName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BaseName");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNodes ChildNodes
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ChildNodes");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLNodes newClass = new LateBindingApi.Office.CustomXMLNodes(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNode FirstChild
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FirstChild");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLNode newClass = new LateBindingApi.Office.CustomXMLNode(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNode LastChild
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LastChild");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLNode newClass = new LateBindingApi.Office.CustomXMLNode(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string NamespaceURI
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NamespaceURI");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNode NextSibling
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NextSibling");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLNode newClass = new LateBindingApi.Office.CustomXMLNode(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoCustomXMLNodeType NodeType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NodeType");
				return (LateBindingApi.Office.Enums.MsoCustomXMLNodeType)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string NodeValue
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NodeValue");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "NodeValue", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMObject OwnerDocument
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OwnerDocument");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLPart OwnerPart
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OwnerPart");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLPart newClass = new LateBindingApi.Office.CustomXMLPart(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNode PreviousSibling
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PreviousSibling");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLNode newClass = new LateBindingApi.Office.CustomXMLNode(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNode ParentNode
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ParentNode");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLNode newClass = new LateBindingApi.Office.CustomXMLNode(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string Text
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Text");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Text", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public string XPath
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "XPath");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string XML
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "XML");
				return (string)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public void AppendChildNode(string name, string namespaceURI, LateBindingApi.Office.Enums.MsoCustomXMLNodeType nodeType, string nodeValue)
		{
			object[] paramArray = new object[4];
			paramArray[0] = name;
			paramArray[1] = namespaceURI;
			paramArray[2] = nodeType;
			paramArray[3] = nodeValue;
			Invoker.Method(this, "AppendChildNode", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void AppendChildSubtree(string xML)
		{
			object[] paramArray = new object[1];
			paramArray[0] = xML;
			Invoker.Method(this, "AppendChildSubtree", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public bool HasChildNodes()
		{
			object returnValue = Invoker.MethodReturn(this, "HasChildNodes", null);
			return (bool)returnValue;
		}

		[SupportByLibrary("OF12","OF14")]
		public void InsertNodeBefore(string name, string namespaceURI, LateBindingApi.Office.Enums.MsoCustomXMLNodeType nodeType, string nodeValue, LateBindingApi.Office.CustomXMLNode nextSibling)
		{
			object[] paramArray = new object[5];
			paramArray[0] = name;
			paramArray[1] = namespaceURI;
			paramArray[2] = nodeType;
			paramArray[3] = nodeValue;
			paramArray.SetValue(nextSibling,4);
			Invoker.Method(this, "InsertNodeBefore", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void InsertSubtreeBefore(string xML, LateBindingApi.Office.CustomXMLNode nextSibling)
		{
			object[] paramArray = new object[2];
			paramArray[0] = xML;
			paramArray.SetValue(nextSibling,1);
			Invoker.Method(this, "InsertSubtreeBefore", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void RemoveChild(LateBindingApi.Office.CustomXMLNode child)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(child,0);
			Invoker.Method(this, "RemoveChild", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void ReplaceChildNode(LateBindingApi.Office.CustomXMLNode oldNode, string name, string namespaceURI, LateBindingApi.Office.Enums.MsoCustomXMLNodeType nodeType, string nodeValue)
		{
			object[] paramArray = new object[5];
			paramArray.SetValue(oldNode,0);
			paramArray[1] = name;
			paramArray[2] = namespaceURI;
			paramArray[3] = nodeType;
			paramArray[4] = nodeValue;
			Invoker.Method(this, "ReplaceChildNode", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void ReplaceChildSubtree(string xML, LateBindingApi.Office.CustomXMLNode oldNode)
		{
			object[] paramArray = new object[2];
			paramArray[0] = xML;
			paramArray.SetValue(oldNode,1);
			Invoker.Method(this, "ReplaceChildSubtree", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNodes SelectNodes(string xPath)
		{
			object[] paramArray = new object[1];
			paramArray[0] = xPath;
			object returnValue = Invoker.MethodReturn(this, "SelectNodes", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CustomXMLNodes newClass = new LateBindingApi.Office.CustomXMLNodes(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLNode SelectSingleNode(string xPath)
		{
			object[] paramArray = new object[1];
			paramArray[0] = xPath;
			object returnValue = Invoker.MethodReturn(this, "SelectSingleNode", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CustomXMLNode newClass = new LateBindingApi.Office.CustomXMLNode(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
