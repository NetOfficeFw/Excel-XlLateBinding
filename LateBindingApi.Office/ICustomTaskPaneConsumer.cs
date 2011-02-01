//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class ICustomTaskPaneConsumer : COMObject
	{
		#region Construction

		public ICustomTaskPaneConsumer(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ICustomTaskPaneConsumer(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ICustomTaskPaneConsumer(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ICustomTaskPaneConsumer()
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public void CTPFactoryAvailable(LateBindingApi.Office.ICTPFactory cTPFactoryInst)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(cTPFactoryInst,0);
			Invoker.Method(this, "CTPFactoryAvailable", paramArray);
		}

		#endregion

	}
}
