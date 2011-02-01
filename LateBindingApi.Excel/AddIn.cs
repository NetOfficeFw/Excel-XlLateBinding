//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class AddIn : COMObject
	{
		#region Construction

		public AddIn(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public AddIn(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public AddIn(COMObject replacedObject) : base(replacedObject)
		{
		}

		public AddIn()
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
		public string Author
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Author");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Comments
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Comments");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string FullName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FullName");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Installed
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Installed");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Installed", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Keywords
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Keywords");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Path
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Path");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Subject
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Subject");
				return (string)returnValue;
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
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string progID
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "progID");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string CLSID
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CLSID");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL14")]
		public bool IsOpen
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsOpen");
				return (bool)returnValue;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
