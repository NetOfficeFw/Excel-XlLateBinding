//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class Validation : COMObject
	{
		#region Construction

		public Validation(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Validation(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Validation(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Validation()
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
		public Int32 AlertStyle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AlertStyle");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IgnoreBlank
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IgnoreBlank");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IgnoreBlank", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 IMEMode
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IMEMode");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IMEMode", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool InCellDropdown
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InCellDropdown");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InCellDropdown", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string ErrorMessage
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ErrorMessage");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ErrorMessage", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string ErrorTitle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ErrorTitle");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ErrorTitle", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string InputMessage
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InputMessage");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InputMessage", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string InputTitle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InputTitle");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InputTitle", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Formula1
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Formula1");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Formula2
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Formula2");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Operator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Operator");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool ShowError
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShowError");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ShowError", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool ShowInput
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShowInput");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ShowInput", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				return (bool)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Add(LateBindingApi.Excel.Enums.XlDVType type)
		{
			object[] paramArray = new object[1];
			paramArray[0] = type;
			Invoker.Method(this, "Add", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Add(LateBindingApi.Excel.Enums.XlDVType type, object alertStyle, object _operator, object formula1, object formula2)
		{
			object[] paramArray = new object[5];
			paramArray[0] = type;
			paramArray[1] = alertStyle;
			paramArray[2] = _operator;
			paramArray[3] = formula1;
			paramArray[4] = formula2;
			Invoker.Method(this, "Add", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Modify()
		{
			Invoker.Method(this, "Modify", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Modify(object type, object alertStyle, object _operator, object formula1, object formula2)
		{
			object[] paramArray = new object[5];
			paramArray[0] = type;
			paramArray[1] = alertStyle;
			paramArray[2] = _operator;
			paramArray[3] = formula1;
			paramArray[4] = formula2;
			Invoker.Method(this, "Modify", paramArray);
		}

		#endregion

	}
}
