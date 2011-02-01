//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class FormatCondition : COMObject
	{
		#region Construction

		public FormatCondition(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public FormatCondition(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public FormatCondition(COMObject replacedObject) : base(replacedObject)
		{
		}

		public FormatCondition()
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
		public Int32 Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (Int32)returnValue;
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
		public LateBindingApi.Excel.Interior Interior
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Interior");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Interior newClass = new LateBindingApi.Excel.Interior(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Borders Borders
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Borders");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Borders newClass = new LateBindingApi.Excel.Borders(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Font Font
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Font");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Font newClass = new LateBindingApi.Excel.Font(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
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


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlContainsOperator TextOperator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextOperator");
				return (LateBindingApi.Excel.Enums.XlContainsOperator)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextOperator", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlTimePeriods DateOperator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DateOperator");
				return (LateBindingApi.Excel.Enums.XlTimePeriods)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DateOperator", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public COMVariant NumberFormat
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NumberFormat");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "NumberFormat", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public Int32 Priority
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Priority");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Priority", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public bool StopIfTrue
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "StopIfTrue");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "StopIfTrue", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Range AppliesTo
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AppliesTo");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public bool PTCondition
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PTCondition");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlPivotConditionScope ScopeType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ScopeType");
				return (LateBindingApi.Excel.Enums.XlPivotConditionScope)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ScopeType", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Modify(LateBindingApi.Excel.Enums.XlFormatConditionType type)
		{
			object[] paramArray = new object[1];
			paramArray[0] = type;
			Invoker.Method(this, "Modify", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL9")]
		public void Modify(LateBindingApi.Excel.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2)
		{
			object[] paramArray = new object[4];
			paramArray[0] = type;
			paramArray[1] = _operator;
			paramArray[2] = formula1;
			paramArray[3] = formula2;
			Invoker.Method(this, "Modify", paramArray);
		}

		[SupportByLibrary("XL12","XL14")]
		public void Modify(LateBindingApi.Excel.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object operator2)
		{
			object[] paramArray = new object[6];
			paramArray[0] = type;
			paramArray[1] = _operator;
			paramArray[2] = formula1;
			paramArray[3] = formula2;
			paramArray[4] = _string;
			paramArray[5] = operator2;
			Invoker.Method(this, "Modify", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("XL12","XL14")]
		public void _Modify(LateBindingApi.Excel.Enums.XlFormatConditionType type)
		{
			object[] paramArray = new object[1];
			paramArray[0] = type;
			Invoker.Method(this, "_Modify", paramArray);
		}

		[SupportByLibrary("XL12","XL14")]
		public void _Modify(LateBindingApi.Excel.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2)
		{
			object[] paramArray = new object[4];
			paramArray[0] = type;
			paramArray[1] = _operator;
			paramArray[2] = formula1;
			paramArray[3] = formula2;
			Invoker.Method(this, "_Modify", paramArray);
		}

		[SupportByLibrary("XL12","XL14")]
		public void ModifyAppliesToRange(LateBindingApi.Excel.Range range)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(range,0);
			Invoker.Method(this, "ModifyAppliesToRange", paramArray);
		}

		[SupportByLibrary("XL12","XL14")]
		public void SetFirstPriority()
		{
			Invoker.Method(this, "SetFirstPriority", null);
		}

		[SupportByLibrary("XL12","XL14")]
		public void SetLastPriority()
		{
			Invoker.Method(this, "SetLastPriority", null);
		}

		#endregion

	}
}