//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class TextBox : COMObject
	{
		#region Construction

		public TextBox(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public TextBox(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public TextBox(COMObject replacedObject) : base(replacedObject)
		{
		}

		public TextBox()
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
		public LateBindingApi.Excel.Range BottomRightCell
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BottomRightCell");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Enabled
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Enabled");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Enabled", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Height
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Height");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Height", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Index
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Index");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Left
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Left");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Left", value);
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
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Name", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string OnAction
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OnAction");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OnAction", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Placement
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Placement");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Placement", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool PrintObject
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintObject");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintObject", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Top
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Top");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Top", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range TopLeftCell
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TopLeftCell");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Visible
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Visible");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Visible", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Width
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Width");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Width", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ZOrder
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ZOrder");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.ShapeRange ShapeRange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShapeRange");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.ShapeRange newClass = new LateBindingApi.Excel.ShapeRange(this, returnValue);
				return newClass;
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
		public COMVariant AutoScaleFont
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AutoScaleFont");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "AutoScaleFont", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool AutoSize
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AutoSize");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AutoSize", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Caption
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Caption");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Caption", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Characters Characters
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Characters", null);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Characters newClass = new LateBindingApi.Excel.Characters(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Characters get_Characters(object start, object length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "Characters", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Characters newClass = new LateBindingApi.Excel.Characters(this, returnValue);
			return newClass;
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
		public string Formula
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Formula");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Formula", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool LockedText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LockedText");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LockedText", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Border Border
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Border");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Border newClass = new LateBindingApi.Excel.Border(this, returnValue);
				return newClass;
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
		public bool RoundedCorners
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RoundedCorners");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RoundedCorners", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Shadow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Shadow");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Shadow", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant BringToFront()
		{
			object returnValue = Invoker.MethodReturn(this, "BringToFront", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Copy()
		{
			object returnValue = Invoker.MethodReturn(this, "Copy", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant CopyPicture(LateBindingApi.Excel.Enums.XlPictureAppearance appearance, LateBindingApi.Excel.Enums.XlCopyPictureFormat format)
		{
			object[] paramArray = new object[2];
			paramArray[0] = appearance;
			paramArray[1] = format;
			object returnValue = Invoker.MethodReturn(this, "CopyPicture", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Cut()
		{
			object returnValue = Invoker.MethodReturn(this, "Cut", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Delete()
		{
			object returnValue = Invoker.MethodReturn(this, "Delete", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Duplicate()
		{
			object returnValue = Invoker.MethodReturn(this, "Duplicate", null);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Select()
		{
			object returnValue = Invoker.MethodReturn(this, "Select", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Select(object replace)
		{
			object[] paramArray = new object[1];
			paramArray[0] = replace;
			object returnValue = Invoker.MethodReturn(this, "Select", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SendToBack()
		{
			object returnValue = Invoker.MethodReturn(this, "SendToBack", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant CheckSpelling()
		{
			object returnValue = Invoker.MethodReturn(this, "CheckSpelling", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object spellLang)
		{
			object[] paramArray = new object[4];
			paramArray[0] = customDictionary;
			paramArray[1] = ignoreUppercase;
			paramArray[2] = alwaysSuggest;
			paramArray[3] = spellLang;
			object returnValue = Invoker.MethodReturn(this, "CheckSpelling", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
