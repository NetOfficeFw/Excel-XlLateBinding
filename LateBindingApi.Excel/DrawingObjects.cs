//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class DrawingObjects : COMObject
	{
		#region Construction

		public DrawingObjects(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public DrawingObjects(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public DrawingObjects(COMObject replacedObject) : base(replacedObject)
		{
		}

		public DrawingObjects()
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
		public COMVariant Accelerator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Accelerator");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Accelerator", value);
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
		public bool CancelButton
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CancelButton");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CancelButton", value);
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
		public bool DefaultButton
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DefaultButton");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DefaultButton", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DismissButton
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DismissButton");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DismissButton", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Display3DShading
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Display3DShading");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Display3DShading", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayVerticalScrollBar
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayVerticalScrollBar");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayVerticalScrollBar", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 DropDownLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DropDownLines");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DropDownLines", value);
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
		public bool HelpButton
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HelpButton");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HelpButton", value);
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
		public Int32 InputType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InputType");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InputType", value);
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
		public Int32 LargeChange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LargeChange");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LargeChange", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string LinkedCell
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LinkedCell");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LinkedCell", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string ListFillRange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListFillRange");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ListFillRange", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ListIndex
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListIndex");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ListIndex", value);
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
		public Int32 Max
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Max");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Max", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Min
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Min");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Min", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool MultiLine
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MultiLine");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "MultiLine", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool MultiSelect
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MultiSelect");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "MultiSelect", value);
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
		public COMVariant PhoneticAccelerator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PhoneticAccelerator");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "PhoneticAccelerator", value);
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
		public Int32 SmallChange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SmallChange");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SmallChange", value);
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
		public Int32 Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Value", value);
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
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy3()
		{
			Invoker.Method(this, "_Dummy3", null);
		}

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
		public void _Dummy12()
		{
			Invoker.Method(this, "_Dummy12", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy15()
		{
			Invoker.Method(this, "_Dummy15", null);
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
		public void _Dummy22()
		{
			Invoker.Method(this, "_Dummy22", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy28()
		{
			Invoker.Method(this, "_Dummy28", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant AddItem(object text)
		{
			object[] paramArray = new object[1];
			paramArray[0] = text;
			object returnValue = Invoker.MethodReturn(this, "AddItem", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant AddItem(object text, object index)
		{
			object[] paramArray = new object[2];
			paramArray[0] = text;
			paramArray[1] = index;
			object returnValue = Invoker.MethodReturn(this, "AddItem", paramArray);
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy47()
		{
			Invoker.Method(this, "_Dummy47", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy54()
		{
			Invoker.Method(this, "_Dummy54", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant List()
		{
			object returnValue = Invoker.MethodReturn(this, "List", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant List(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.MethodReturn(this, "List", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void _Dummy56()
		{
			Invoker.Method(this, "_Dummy56", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant RemoveAllItems()
		{
			object returnValue = Invoker.MethodReturn(this, "RemoveAllItems", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant RemoveItem(Int32 index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.MethodReturn(this, "RemoveItem", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant RemoveItem(Int32 index, object count)
		{
			object[] paramArray = new object[2];
			paramArray[0] = index;
			paramArray[1] = count;
			object returnValue = Invoker.MethodReturn(this, "RemoveItem", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Reshape(Int32 vertex, object insert)
		{
			object[] paramArray = new object[2];
			paramArray[0] = vertex;
			paramArray[1] = insert;
			object returnValue = Invoker.MethodReturn(this, "Reshape", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Reshape(Int32 vertex, object insert, object left, object top)
		{
			object[] paramArray = new object[4];
			paramArray[0] = vertex;
			paramArray[1] = insert;
			paramArray[2] = left;
			paramArray[3] = top;
			object returnValue = Invoker.MethodReturn(this, "Reshape", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Selected()
		{
			object returnValue = Invoker.MethodReturn(this, "Selected", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Selected(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.MethodReturn(this, "Selected", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Ungroup()
		{
			object returnValue = Invoker.MethodReturn(this, "Ungroup", null);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Vertices()
		{
			object returnValue = Invoker.MethodReturn(this, "Vertices", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Vertices(object index1, object index2)
		{
			object[] paramArray = new object[2];
			paramArray[0] = index1;
			paramArray[1] = index2;
			object returnValue = Invoker.MethodReturn(this, "Vertices", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.GroupObject Group()
		{
			object returnValue = Invoker.MethodReturn(this, "Group", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.GroupObject newClass = new LateBindingApi.Excel.GroupObject(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant LinkCombo()
		{
			object returnValue = Invoker.MethodReturn(this, "LinkCombo", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant LinkCombo(object link)
		{
			object[] paramArray = new object[1];
			paramArray[0] = link;
			object returnValue = Invoker.MethodReturn(this, "LinkCombo", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.MethodReturn(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				COMObject returnClass = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

	}
}
