//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class ICTPFactory : COMObject
	{
		#region Construction

		public ICTPFactory(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ICTPFactory(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ICTPFactory(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ICTPFactory()
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle)
		{
			object[] paramArray = new object[2];
			paramArray[0] = cTPAxID;
			paramArray[1] = cTPTitle;
			object returnValue = Invoker.MethodReturn(this, "CreateCTP", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office._CustomTaskPane newClass = new LateBindingApi.Office._CustomTaskPane(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle, object cTPParentWindow)
		{
			object[] paramArray = new object[3];
			paramArray[0] = cTPAxID;
			paramArray[1] = cTPTitle;
			paramArray[2] = cTPParentWindow;
			object returnValue = Invoker.MethodReturn(this, "CreateCTP", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office._CustomTaskPane newClass = new LateBindingApi.Office._CustomTaskPane(this, returnValue);
			return newClass;
		}

		#endregion

	}
}