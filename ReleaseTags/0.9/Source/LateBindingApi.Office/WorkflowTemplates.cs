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
	public class WorkflowTemplates : _IMsoDispObj
	{
		#region Construction

		public WorkflowTemplates(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public WorkflowTemplates(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public WorkflowTemplates(COMObject replacedObject) : base(replacedObject)
		{
		}

		public WorkflowTemplates()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.WorkflowTemplate this[Int32 index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.WorkflowTemplate newClass = new LateBindingApi.Office.WorkflowTemplate(this, returnValue);
				return newClass;
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

		[SupportByLibrary("OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.WorkflowTemplate returnClass = new LateBindingApi.Office.WorkflowTemplate (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		#endregion

	}
}