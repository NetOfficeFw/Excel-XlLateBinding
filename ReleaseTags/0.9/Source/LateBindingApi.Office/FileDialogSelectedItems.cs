//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class FileDialogSelectedItems : _IMsoDispObj
	{
		#region Construction

		public FileDialogSelectedItems(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public FileDialogSelectedItems(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public FileDialogSelectedItems(COMObject replacedObject) : base(replacedObject)
		{
		}

		public FileDialogSelectedItems()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				string returnType = (string)itemProxy;
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnType;
            }
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
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

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string this[Int32 index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				return (string)returnValue;
			}
		}

		#endregion

	}
}
