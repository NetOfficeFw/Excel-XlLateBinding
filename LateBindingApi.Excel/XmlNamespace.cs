//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL11","XL12","XL14")]
	public class XmlNamespace : COMObject
	{
		#region Construction

		public XmlNamespace(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public XmlNamespace(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public XmlNamespace(COMObject replacedObject) : base(replacedObject)
		{
		}

		public XmlNamespace()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.Application Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Application newClass = new LateBindingApi.Excel.Application(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public string _Default
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_Default");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public string Uri
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Uri");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public string Prefix
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Prefix");
				return (string)returnValue;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
