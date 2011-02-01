//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class _dispVBProjectsEvents : COMObject
	{
		#region Construction

		public _dispVBProjectsEvents(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _dispVBProjectsEvents(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _dispVBProjectsEvents(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _dispVBProjectsEvents()
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		[SupportByLibrary("VBE")]
		public void ItemAdded(LateBindingApi.VBIDE.VBProject vBProject)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(vBProject,0);
			Invoker.Method(this, "ItemAdded", paramArray);
		}

		[SupportByLibrary("VBE")]
		public void ItemRemoved(LateBindingApi.VBIDE.VBProject vBProject)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(vBProject,0);
			Invoker.Method(this, "ItemRemoved", paramArray);
		}

		[SupportByLibrary("VBE")]
		public void ItemRenamed(LateBindingApi.VBIDE.VBProject vBProject, string oldName)
		{
			object[] paramArray = new object[2];
			paramArray.SetValue(vBProject,0);
			paramArray[1] = oldName;
			Invoker.Method(this, "ItemRenamed", paramArray);
		}

		[SupportByLibrary("VBE")]
		public void ItemActivated(LateBindingApi.VBIDE.VBProject vBProject)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(vBProject,0);
			Invoker.Method(this, "ItemActivated", paramArray);
		}

		#endregion

	}
}
