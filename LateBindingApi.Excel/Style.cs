//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class Style : COMObject
	{
		#region Construction

		public Style(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Style(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Style(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Style()
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
		public bool AddIndent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AddIndent");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AddIndent", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool BuiltIn
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BuiltIn");
				return (bool)returnValue;
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool FormulaHidden
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FormulaHidden");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FormulaHidden", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlHAlign HorizontalAlignment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HorizontalAlignment");
				return (LateBindingApi.Excel.Enums.XlHAlign)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HorizontalAlignment", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IncludeAlignment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IncludeAlignment");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IncludeAlignment", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IncludeBorder
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IncludeBorder");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IncludeBorder", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IncludeFont
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IncludeFont");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IncludeFont", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IncludeNumber
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IncludeNumber");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IncludeNumber", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IncludePatterns
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IncludePatterns");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IncludePatterns", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IncludeProtection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IncludeProtection");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IncludeProtection", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 IndentLevel
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IndentLevel");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "IndentLevel", value);
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
		public bool Locked
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Locked");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Locked", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant MergeCells
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MergeCells");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "MergeCells", value);
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
		public string NameLocal
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NameLocal");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string NumberFormat
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NumberFormat");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "NumberFormat", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string NumberFormatLocal
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NumberFormatLocal");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "NumberFormatLocal", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlOrientation Orientation
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Orientation");
				return (LateBindingApi.Excel.Enums.XlOrientation)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Orientation", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool ShrinkToFit
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShrinkToFit");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ShrinkToFit", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlVAlign VerticalAlignment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VerticalAlignment");
				return (LateBindingApi.Excel.Enums.XlVAlign)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "VerticalAlignment", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool WrapText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WrapText");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WrapText", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string _Default
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_Default");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ReadingOrder
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ReadingOrder");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ReadingOrder", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Delete()
		{
			object returnValue = Invoker.MethodReturn(this, "Delete", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}