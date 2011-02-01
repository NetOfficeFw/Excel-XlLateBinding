//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class FreeformBuilder : COMObject
	{
		#region Construction

		public FreeformBuilder(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public FreeformBuilder(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public FreeformBuilder(COMObject replacedObject) : base(replacedObject)
		{
		}

		public FreeformBuilder()
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

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void AddNodes(LateBindingApi.Office.Enums.MsoSegmentType segmentType, LateBindingApi.Office.Enums.MsoEditingType editingType, Double x1, Double y1)
		{
			object[] paramArray = new object[4];
			paramArray[0] = segmentType;
			paramArray[1] = editingType;
			paramArray[2] = x1;
			paramArray[3] = y1;
			Invoker.Method(this, "AddNodes", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void AddNodes(LateBindingApi.Office.Enums.MsoSegmentType segmentType, LateBindingApi.Office.Enums.MsoEditingType editingType, Double x1, Double y1, object x2, object y2, object x3, object y3)
		{
			object[] paramArray = new object[8];
			paramArray[0] = segmentType;
			paramArray[1] = editingType;
			paramArray[2] = x1;
			paramArray[3] = y1;
			paramArray[4] = x2;
			paramArray[5] = y2;
			paramArray[6] = x3;
			paramArray[7] = y3;
			Invoker.Method(this, "AddNodes", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Shape ConvertToShape()
		{
			object returnValue = Invoker.MethodReturn(this, "ConvertToShape", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Shape newClass = new LateBindingApi.Excel.Shape(this, returnValue);
			return newClass;
		}

		#endregion

	}
}