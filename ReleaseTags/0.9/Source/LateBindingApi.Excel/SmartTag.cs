//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14")]
	public class SmartTag : COMObject
	{
		#region Construction

		public SmartTag(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SmartTag(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SmartTag(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SmartTag()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
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

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string DownloadURL
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DownloadURL");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string _Default
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_Default");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string XML
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "XML");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Range Range
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Range");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.SmartTagActions SmartTagActions
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SmartTagActions");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.SmartTagActions newClass = new LateBindingApi.Excel.SmartTagActions(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.CustomProperties Properties
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Properties");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.CustomProperties newClass = new LateBindingApi.Excel.CustomProperties(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		#endregion

	}
}
