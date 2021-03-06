//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class CustomXMLPrefixMapping : _IMsoDispObj
	{
		#region Construction

		public CustomXMLPrefixMapping(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public CustomXMLPrefixMapping(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public CustomXMLPrefixMapping(COMObject replacedObject) : base(replacedObject)
		{
		}

		public CustomXMLPrefixMapping()
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
		public string Prefix
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Prefix");
				return (string)returnValue;
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

		#endregion

		#region Methods

		#endregion

	}
}
