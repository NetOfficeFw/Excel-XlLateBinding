//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class DropDown : COMObject
	{
		#region Construction

		public DropDown(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public DropDown(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public DropDown(COMObject replacedObject) : base(replacedObject)
		{
		}

		public DropDown()
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
		public COMVariant LinkedObject
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LinkedObject");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant List
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "List", null);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "List", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant get_List(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
				object returnValue = Invoker.PropertyGet(this, "List", paramArray);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
		}
		
		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void set_List(object index, object value)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.PropertySet(this, "List", paramArray, value);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ListCount
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListCount");
				return (Int32)returnValue;
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
		public COMVariant Selected
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Selected", null);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Selected", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant get_Selected(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
				object returnValue = Invoker.PropertyGet(this, "Selected", paramArray);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
		}
		
		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void set_Selected(object index, object value)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.PropertySet(this, "Selected", paramArray, value);
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
		public void _Dummy36()
		{
			Invoker.Method(this, "_Dummy36", null);
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

		#endregion

	}
}
