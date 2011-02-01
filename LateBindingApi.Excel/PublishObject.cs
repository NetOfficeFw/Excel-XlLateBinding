//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class PublishObject : COMObject
	{
		#region Construction

		public PublishObject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PublishObject(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PublishObject(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PublishObject()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string DivID
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DivID");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Sheet
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Sheet");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlSourceType SourceType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceType");
				return (LateBindingApi.Excel.Enums.XlSourceType)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Source
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Source");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlHtmlType HtmlType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HtmlType");
				return (LateBindingApi.Excel.Enums.XlHtmlType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HtmlType", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Title
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Title");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Title", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Filename
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Filename");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Filename", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public bool AutoRepublish
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AutoRepublish");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AutoRepublish", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Publish()
		{
			Invoker.Method(this, "Publish", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Publish(object create)
		{
			object[] paramArray = new object[1];
			paramArray[0] = create;
			Invoker.Method(this, "Publish", paramArray);
		}

		#endregion

	}
}