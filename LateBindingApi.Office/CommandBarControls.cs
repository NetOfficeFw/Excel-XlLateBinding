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
	public class CommandBarControls : _IMsoDispObj
	{
		#region Construction

		public CommandBarControls(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public CommandBarControls(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public CommandBarControls(COMObject replacedObject) : base(replacedObject)
		{
		}

		public CommandBarControls()
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
		public LateBindingApi.Office.CommandBarControl this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CommandBarControl newClass = new LateBindingApi.Office.CommandBarControl(this, returnValue);
				return newClass;
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
				LateBindingApi.Office.CommandBarControl returnClass = new LateBindingApi.Office.CommandBarControl (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBar Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CommandBar newClass = new LateBindingApi.Office.CommandBar(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBarControl Add()
		{
			object returnValue = Invoker.MethodReturn(this, "Add", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CommandBarControl newClass = new LateBindingApi.Office.CommandBarControl(this, returnValue);
			return newClass;
		}

        [SupportByLibrary( "OF10", "OF11", "OF12", "OF14","OF9")]
        public LateBindingApi.Office.CommandBarControl Add(object type)
        {
            object[] paramArray = new object[1];
            paramArray[0] = type;
            object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
            COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
            return (LateBindingApi.Office.CommandBarControl)returnObject;
        }

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBarControl Add(object type, object id, object parameter, object before, object temporary)
		{
			object[] paramArray = new object[5];
			paramArray[0] = type;
			paramArray[1] = id;
			paramArray[2] = parameter;
			paramArray[3] = before;
			paramArray[4] = temporary;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
            COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
            return (LateBindingApi.Office.CommandBarControl)returnObject;
		}

		#endregion

	}
}