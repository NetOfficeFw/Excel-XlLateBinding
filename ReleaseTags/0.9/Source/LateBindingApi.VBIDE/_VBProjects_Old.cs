//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class _VBProjects_Old : COMObject
	{
		#region Construction

		public _VBProjects_Old(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _VBProjects_Old(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _VBProjects_Old(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _VBProjects_Old()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.VBE VBE
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VBE");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.VBE newClass = new LateBindingApi.VBIDE.VBE(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.VBE Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.VBE newClass = new LateBindingApi.VBIDE.VBE(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
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

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.VBProject this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.VBProject newClass = new LateBindingApi.VBIDE.VBProject(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.MethodReturn(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.VBIDE.VBProject returnClass = new LateBindingApi.VBIDE.VBProject (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

	}
}