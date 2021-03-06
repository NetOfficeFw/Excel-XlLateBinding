//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF14")]
	public class SmartArtNodes : _IMsoDispObj
	{
		#region Construction

		public SmartArtNodes(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SmartArtNodes(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SmartArtNodes(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SmartArtNodes()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.SmartArtNode returnClass = new LateBindingApi.Office.SmartArtNode (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF14")]
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

		[SupportByLibrary("OF14")]
		public LateBindingApi.Office.SmartArtNode this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.SmartArtNode newClass = new LateBindingApi.Office.SmartArtNode(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF14")]
		public LateBindingApi.Office.SmartArtNode Add()
		{
			object returnValue = Invoker.MethodReturn(this, "Add", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.SmartArtNode newClass = new LateBindingApi.Office.SmartArtNode(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
