//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class Pane : COMObject
	{
		#region Construction

		public Pane(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Pane(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Pane(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Pane()
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
		public Int32 Index
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Index");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ScrollColumn
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ScrollColumn");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ScrollColumn", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ScrollRow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ScrollRow");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ScrollRow", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range VisibleRange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VisibleRange");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Activate()
		{
			object returnValue = Invoker.MethodReturn(this, "Activate", null);
			return (bool)returnValue;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant LargeScroll()
		{
			object returnValue = Invoker.MethodReturn(this, "LargeScroll", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant LargeScroll(object down, object up, object toRight, object toLeft)
		{
			object[] paramArray = new object[4];
			paramArray[0] = down;
			paramArray[1] = up;
			paramArray[2] = toRight;
			paramArray[3] = toLeft;
			object returnValue = Invoker.MethodReturn(this, "LargeScroll", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SmallScroll()
		{
			object returnValue = Invoker.MethodReturn(this, "SmallScroll", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SmallScroll(object down, object up, object toRight, object toLeft)
		{
			object[] paramArray = new object[4];
			paramArray[0] = down;
			paramArray[1] = up;
			paramArray[2] = toRight;
			paramArray[3] = toLeft;
			object returnValue = Invoker.MethodReturn(this, "SmallScroll", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void ScrollIntoView(Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramArray = new object[4];
			paramArray[0] = left;
			paramArray[1] = top;
			paramArray[2] = width;
			paramArray[3] = height;
			Invoker.Method(this, "ScrollIntoView", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void ScrollIntoView(Int32 left, Int32 top, Int32 width, Int32 height, object start)
		{
			object[] paramArray = new object[5];
			paramArray[0] = left;
			paramArray[1] = top;
			paramArray[2] = width;
			paramArray[3] = height;
			paramArray[4] = start;
			Invoker.Method(this, "ScrollIntoView", paramArray);
		}

		[SupportByLibrary("XL12","XL14")]
		public Int32 PointsToScreenPixelsX(Int32 points)
		{
			object[] paramArray = new object[1];
			paramArray[0] = points;
			object returnValue = Invoker.MethodReturn(this, "PointsToScreenPixelsX", paramArray);
			return (Int32)returnValue;
		}

		[SupportByLibrary("XL12","XL14")]
		public Int32 PointsToScreenPixelsY(Int32 points)
		{
			object[] paramArray = new object[1];
			paramArray[0] = points;
			object returnValue = Invoker.MethodReturn(this, "PointsToScreenPixelsY", paramArray);
			return (Int32)returnValue;
		}

		#endregion

	}
}
