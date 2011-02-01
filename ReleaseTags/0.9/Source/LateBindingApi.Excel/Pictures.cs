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
	public class Pictures : COMObject
	{
		#region Construction

		public Pictures(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Pictures(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Pictures(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Pictures()
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
		public LateBindingApi.Excel.Picture Add(Double left, Double top, Double width, Double height)
		{
			object[] paramArray = new object[4];
			paramArray[0] = left;
			paramArray[1] = top;
			paramArray[2] = width;
			paramArray[3] = height;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Picture newClass = new LateBindingApi.Excel.Picture(this, returnValue);
			return newClass;
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
		public LateBindingApi.Excel.Picture Insert(string filename)
		{
			object[] paramArray = new object[1];
			paramArray[0] = filename;
			object returnValue = Invoker.MethodReturn(this, "Insert", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Picture newClass = new LateBindingApi.Excel.Picture(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Picture Insert(string filename, object converter)
		{
			object[] paramArray = new object[2];
			paramArray[0] = filename;
			paramArray[1] = converter;
			object returnValue = Invoker.MethodReturn(this, "Insert", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Picture newClass = new LateBindingApi.Excel.Picture(this, returnValue);
			return newClass;
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Picture Paste()
		{
			object returnValue = Invoker.MethodReturn(this, "Paste", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Picture newClass = new LateBindingApi.Excel.Picture(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Picture Paste(object link)
		{
			object[] paramArray = new object[1];
			paramArray[0] = link;
			object returnValue = Invoker.MethodReturn(this, "Paste", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Picture newClass = new LateBindingApi.Excel.Picture(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
