//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class ChartGroups : COMObject
	{
		#region Construction

		public ChartGroups(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ChartGroups(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ChartGroups(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ChartGroups()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF14")]
		public COMObject Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF14")]
		public Int32 Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (Int32)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.IMsoChartGroup this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IMsoChartGroup newClass = new LateBindingApi.Office.IMsoChartGroup(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.MethodReturn(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.IMsoChartGroup returnClass = new LateBindingApi.Office.IMsoChartGroup (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

	}
}
