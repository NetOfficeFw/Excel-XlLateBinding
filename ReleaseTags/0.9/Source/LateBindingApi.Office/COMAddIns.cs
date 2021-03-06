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
	public class COMAddIns : _IMsoDispObj
	{
		#region Construction

		public COMAddIns(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public COMAddIns(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public COMAddIns(COMObject replacedObject) : base(replacedObject)
		{
		}

		public COMAddIns()
		{
		}

		#endregion

		#region Properties

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
				LateBindingApi.Office.COMAddIn returnClass = new LateBindingApi.Office.COMAddIn (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
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

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.COMAddIn this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.COMAddIn newClass = new LateBindingApi.Office.COMAddIn(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Update()
		{
			Invoker.Method(this, "Update", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void SetAppModal(bool varfModal)
		{
			object[] paramArray = new object[1];
			paramArray[0] = varfModal;
			Invoker.Method(this, "SetAppModal", paramArray);
		}

		#endregion

	}
}
