//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class PropertyTests : _IMsoDispObj
	{
		#region Construction

		public PropertyTests(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PropertyTests(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PropertyTests(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PropertyTests()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.PropertyTest this[Int32 index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.PropertyTest newClass = new LateBindingApi.Office.PropertyTest(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.PropertyTest returnClass = new LateBindingApi.Office.PropertyTest (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Add(string name, LateBindingApi.Office.Enums.MsoCondition condition, object value, object secondValue, LateBindingApi.Office.Enums.MsoConnector connector)
		{
			object[] paramArray = new object[5];
			paramArray[0] = name;
			paramArray[1] = condition;
			paramArray[2] = value;
			paramArray[3] = secondValue;
			paramArray[4] = connector;
			Invoker.Method(this, "Add", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Remove(Int32 index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.Method(this, "Remove", paramArray);
		}

		#endregion

	}
}