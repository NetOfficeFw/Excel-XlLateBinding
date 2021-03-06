//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class GroupObject : COMObject
	{
		#region Construction

		public GroupObject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public GroupObject(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public GroupObject(COMObject replacedObject) : base(replacedObject)
		{
		}

		public GroupObject()
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
		public COMVariant ArrowHeadLength
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ArrowHeadLength");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "ArrowHeadLength", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ArrowHeadStyle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ArrowHeadStyle");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "ArrowHeadStyle", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ArrowHeadWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ArrowHeadWidth");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "ArrowHeadWidth", value);
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
		public Int32 _Default
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_Default");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "_Default", value);
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
		public void _Dummy27()
		{
			Invoker.Method(this, "_Dummy27", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy28()
		{
			Invoker.Method(this, "_Dummy28", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy30()
		{
			Invoker.Method(this, "_Dummy30", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy36()
		{
			Invoker.Method(this, "_Dummy36", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy37()
		{
			Invoker.Method(this, "_Dummy37", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy38()
		{
			Invoker.Method(this, "_Dummy38", null);
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy41()
		{
			Invoker.Method(this, "_Dummy41", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy42()
		{
			Invoker.Method(this, "_Dummy42", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy43()
		{
			Invoker.Method(this, "_Dummy43", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy44()
		{
			Invoker.Method(this, "_Dummy44", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy45()
		{
			Invoker.Method(this, "_Dummy45", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy47()
		{
			Invoker.Method(this, "_Dummy47", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy48()
		{
			Invoker.Method(this, "_Dummy48", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy50()
		{
			Invoker.Method(this, "_Dummy50", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy52()
		{
			Invoker.Method(this, "_Dummy52", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy53()
		{
			Invoker.Method(this, "_Dummy53", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy54()
		{
			Invoker.Method(this, "_Dummy54", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy55()
		{
			Invoker.Method(this, "_Dummy55", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy56()
		{
			Invoker.Method(this, "_Dummy56", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy57()
		{
			Invoker.Method(this, "_Dummy57", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy58()
		{
			Invoker.Method(this, "_Dummy58", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy59()
		{
			Invoker.Method(this, "_Dummy59", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy60()
		{
			Invoker.Method(this, "_Dummy60", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy61()
		{
			Invoker.Method(this, "_Dummy61", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy62()
		{
			Invoker.Method(this, "_Dummy62", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy63()
		{
			Invoker.Method(this, "_Dummy63", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy65()
		{
			Invoker.Method(this, "_Dummy65", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy66()
		{
			Invoker.Method(this, "_Dummy66", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy67()
		{
			Invoker.Method(this, "_Dummy67", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy68()
		{
			Invoker.Method(this, "_Dummy68", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy70()
		{
			Invoker.Method(this, "_Dummy70", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy72()
		{
			Invoker.Method(this, "_Dummy72", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy73()
		{
			Invoker.Method(this, "_Dummy73", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Ungroup()
		{
			object returnValue = Invoker.MethodReturn(this, "Ungroup", null);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy75()
		{
			Invoker.Method(this, "_Dummy75", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy77()
		{
			Invoker.Method(this, "_Dummy77", null);
		}

		#endregion

	}
}
