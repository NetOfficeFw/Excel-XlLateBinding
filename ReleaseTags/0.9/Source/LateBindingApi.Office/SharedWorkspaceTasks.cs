//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF11","OF12","OF14")]
	public class SharedWorkspaceTasks : _IMsoDispObj
	{
		#region Construction

		public SharedWorkspaceTasks(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SharedWorkspaceTasks(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SharedWorkspaceTasks(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SharedWorkspaceTasks()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.SharedWorkspaceTask this[Int32 index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.SharedWorkspaceTask newClass = new LateBindingApi.Office.SharedWorkspaceTask(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public bool ItemCountExceeded
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ItemCountExceeded");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.SharedWorkspaceTask returnClass = new LateBindingApi.Office.SharedWorkspaceTask (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.SharedWorkspaceTask Add(string title)
		{
			object[] paramArray = new object[1];
			paramArray[0] = title;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.SharedWorkspaceTask newClass = new LateBindingApi.Office.SharedWorkspaceTask(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.SharedWorkspaceTask Add(string title, object status, object priority, object assignee, object description, object dueDate)
		{
			object[] paramArray = new object[6];
			paramArray[0] = title;
			paramArray[1] = status;
			paramArray[2] = priority;
			paramArray[3] = assignee;
			paramArray[4] = description;
			paramArray[5] = dueDate;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.SharedWorkspaceTask newClass = new LateBindingApi.Office.SharedWorkspaceTask(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
