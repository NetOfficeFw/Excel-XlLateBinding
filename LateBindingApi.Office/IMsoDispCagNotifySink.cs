//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class IMsoDispCagNotifySink : COMObject
	{
		#region Construction

		public IMsoDispCagNotifySink(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IMsoDispCagNotifySink(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IMsoDispCagNotifySink(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IMsoDispCagNotifySink()
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void InsertClip(object pClipMoniker, object pItemMoniker)
		{
			object[] paramArray = new object[2];
			paramArray[0] = pClipMoniker;
			paramArray[1] = pItemMoniker;
			Invoker.Method(this, "InsertClip", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void WindowIsClosing()
		{
			Invoker.Method(this, "WindowIsClosing", null);
		}

		#endregion

	}
}
