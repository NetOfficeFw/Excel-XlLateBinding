//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL11","XL12","XL14")]
	public class XPath : COMObject
	{
		#region Construction

		public XPath(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public XPath(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public XPath(COMObject replacedObject) : base(replacedObject)
		{
		}

		public XPath()
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
		public string Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.XmlMap Map
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Map");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.XmlMap newClass = new LateBindingApi.Excel.XmlMap(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public bool Repeating
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Repeating");
				return (bool)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL11","XL12","XL14")]
		public void SetValue(LateBindingApi.Excel.XmlMap map, string xPath)
		{
			object[] paramArray = new object[2];
			paramArray.SetValue(map,0);
			paramArray[1] = xPath;
			Invoker.Method(this, "SetValue", paramArray);
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public void SetValue(LateBindingApi.Excel.XmlMap map, string xPath, object selectionNamespace, object repeating)
		{
			object[] paramArray = new object[4];
			paramArray.SetValue(map,0);
			paramArray[1] = xPath;
			paramArray[2] = selectionNamespace;
			paramArray[3] = repeating;
			Invoker.Method(this, "SetValue", paramArray);
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public void Clear()
		{
			Invoker.Method(this, "Clear", null);
		}

		#endregion

	}
}
