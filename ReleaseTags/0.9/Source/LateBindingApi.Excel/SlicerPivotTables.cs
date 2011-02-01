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
	public class SlicerPivotTables : COMObject
	{
		#region Construction

		public SlicerPivotTables(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SlicerPivotTables(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SlicerPivotTables(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SlicerPivotTables()
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
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.PivotTable this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.PivotTable newClass = new LateBindingApi.Excel.PivotTable(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.PivotTable get__Default(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.PivotTable newClass = new LateBindingApi.Excel.PivotTable(this, returnValue);
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
				LateBindingApi.Excel.PivotTable returnClass = new LateBindingApi.Excel.PivotTable (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL14")]
		public void AddPivotTable(LateBindingApi.Excel.PivotTable pivotTable)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(pivotTable,0);
			Invoker.Method(this, "AddPivotTable", paramArray);
		}

		[SupportByLibrary("XL14")]
		public void RemovePivotTable(object pivotTable)
		{
			object[] paramArray = new object[1];
			paramArray[0] = pivotTable;
			Invoker.Method(this, "RemovePivotTable", paramArray);
		}

		#endregion

	}
}
