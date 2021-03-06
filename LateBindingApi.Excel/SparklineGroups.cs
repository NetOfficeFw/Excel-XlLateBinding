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
	public class SparklineGroups : COMObject
	{
		#region Construction

		public SparklineGroups(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SparklineGroups(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SparklineGroups(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SparklineGroups()
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
		public LateBindingApi.Excel.SparklineGroup this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.SparklineGroup newClass = new LateBindingApi.Excel.SparklineGroup(this, returnValue);
				return newClass;
			}
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
				LateBindingApi.Excel.SparklineGroup returnClass = new LateBindingApi.Excel.SparklineGroup (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.SparklineGroup get__Default(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.SparklineGroup newClass = new LateBindingApi.Excel.SparklineGroup(this, returnValue);
			return newClass;
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.SparklineGroup Add(LateBindingApi.Excel.Enums.XlSparkType type, string sourceData)
		{
			object[] paramArray = new object[2];
			paramArray[0] = type;
			paramArray[1] = sourceData;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.SparklineGroup newClass = new LateBindingApi.Excel.SparklineGroup(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL14")]
		public void Clear()
		{
			Invoker.Method(this, "Clear", null);
		}

		[SupportByLibrary("XL14")]
		public void ClearGroups()
		{
			Invoker.Method(this, "ClearGroups", null);
		}

		[SupportByLibrary("XL14")]
		public void Group(LateBindingApi.Excel.Range location)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(location,0);
			Invoker.Method(this, "Group", paramArray);
		}

		[SupportByLibrary("XL14")]
		public void Ungroup()
		{
			Invoker.Method(this, "Ungroup", null);
		}

		#endregion

	}
}
