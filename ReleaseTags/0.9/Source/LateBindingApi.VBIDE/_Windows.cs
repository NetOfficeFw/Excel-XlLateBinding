//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class _Windows : _Windows_old
	{
		#region Construction

		public _Windows(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _Windows(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _Windows(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _Windows()
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Window CreateToolWindow(LateBindingApi.VBIDE.AddIn addInInst, string progId, string caption, string guidPosition, COMObject docObj)
		{
			object[] paramArray = new object[5];
			paramArray.SetValue(addInInst,0);
			paramArray[1] = progId;
			paramArray[2] = caption;
			paramArray[3] = guidPosition;
			paramArray.SetValue(docObj,4);
			object returnValue = Invoker.MethodReturn(this, "CreateToolWindow", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.VBIDE.Window newClass = new LateBindingApi.VBIDE.Window(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
