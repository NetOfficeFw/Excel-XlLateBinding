//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL14")]
	public class PivotTableChangeList : COMObject
	{
		#region Construction

		public PivotTableChangeList(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PivotTableChangeList(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PivotTableChangeList(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PivotTableChangeList()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL14")]
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

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.ValueChange get__Default(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ValueChange newClass = new LateBindingApi.Excel.ValueChange(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Excel.ValueChange returnClass = new LateBindingApi.Excel.ValueChange (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.ValueChange this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.ValueChange newClass = new LateBindingApi.Excel.ValueChange(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL14")]
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

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.ValueChange Add(string tuple, Double value)
		{
			object[] paramArray = new object[2];
			paramArray[0] = tuple;
			paramArray[1] = value;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ValueChange newClass = new LateBindingApi.Excel.ValueChange(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.ValueChange Add(string tuple, Double value, object allocationValue, object allocationMethod, object allocationWeightExpression)
		{
			object[] paramArray = new object[5];
			paramArray[0] = tuple;
			paramArray[1] = value;
			paramArray[2] = allocationValue;
			paramArray[3] = allocationMethod;
			paramArray[4] = allocationWeightExpression;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ValueChange newClass = new LateBindingApi.Excel.ValueChange(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
