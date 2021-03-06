//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14")]
	public class CellFormat : COMObject
	{
		#region Construction

		public CellFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public CellFormat(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public CellFormat(COMObject replacedObject) : base(replacedObject)
		{
		}

		public CellFormat()
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
			set
			{
				Invoker.PropertySet(this, "Borders", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
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
			set
			{
				Invoker.PropertySet(this, "Font", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
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
			set
			{
				Invoker.PropertySet(this, "Interior", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant NumberFormatLocal
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NumberFormatLocal");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "NumberFormatLocal", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant AddIndent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AddIndent");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "AddIndent", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant IndentLevel
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IndentLevel");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "IndentLevel", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant HorizontalAlignment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HorizontalAlignment");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "HorizontalAlignment", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant VerticalAlignment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VerticalAlignment");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "VerticalAlignment", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant Orientation
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Orientation");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Orientation", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant ShrinkToFit
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShrinkToFit");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "ShrinkToFit", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant WrapText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WrapText");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "WrapText", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant Locked
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Locked");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Locked", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant FormulaHidden
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FormulaHidden");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "FormulaHidden", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
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


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void Clear()
		{
			Invoker.Method(this, "Clear", null);
		}

		#endregion

	}
}
