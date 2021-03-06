//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14")]
	public class SmartTagAction : COMObject
	{
		#region Construction

		public SmartTagAction(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SmartTagAction(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SmartTagAction(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SmartTagAction()
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

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlSmartTagControlType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Excel.Enums.XlSmartTagControlType)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public bool PresentInPane
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PresentInPane");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public bool ExpandHelp
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ExpandHelp");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ExpandHelp", value);
			}						
		}


		[SupportByLibrary("XL11","XL12","XL14")]
		public bool CheckboxState
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CheckboxState");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CheckboxState", value);
			}						
		}


		[SupportByLibrary("XL11","XL12","XL14")]
		public string TextboxText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextboxText");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextboxText", value);
			}						
		}


		[SupportByLibrary("XL11","XL12","XL14")]
		public Int32 ListSelection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListSelection");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ListSelection", value);
			}						
		}


		[SupportByLibrary("XL11","XL12","XL14")]
		public Int32 RadioGroupSelection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RadioGroupSelection");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RadioGroupSelection", value);
			}						
		}


		[SupportByLibrary("XL11","XL12","XL14")]
		public COMObject ActiveXControl
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ActiveXControl");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void Execute()
		{
			Invoker.Method(this, "Execute", null);
		}

		#endregion

	}
}
