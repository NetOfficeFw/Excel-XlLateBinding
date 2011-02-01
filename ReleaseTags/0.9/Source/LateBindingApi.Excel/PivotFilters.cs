//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL12","XL14")]
	public class PivotFilters : COMObject
	{
		#region Construction

		public PivotFilters(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PivotFilters(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PivotFilters(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PivotFilters()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL12","XL14")]
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

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.PivotFilter get__Default(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.PivotFilter newClass = new LateBindingApi.Excel.PivotFilter(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL12","XL14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Excel.PivotFilter returnClass = new LateBindingApi.Excel.PivotFilter (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.PivotFilter this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.PivotFilter newClass = new LateBindingApi.Excel.PivotFilter(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
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

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.PivotFilter Add(LateBindingApi.Excel.Enums.XlPivotFilterType type)
		{
			object[] paramArray = new object[1];
			paramArray[0] = type;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.PivotFilter newClass = new LateBindingApi.Excel.PivotFilter(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.PivotFilter Add(LateBindingApi.Excel.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField)
		{
			object[] paramArray = new object[8];
			paramArray[0] = type;
			paramArray[1] = dataField;
			paramArray[2] = value1;
			paramArray[3] = value2;
			paramArray[4] = order;
			paramArray[5] = name;
			paramArray[6] = description;
			paramArray[7] = memberPropertyField;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.PivotFilter newClass = new LateBindingApi.Excel.PivotFilter(this, returnValue);
			return newClass;
		}

		#endregion

	}
}