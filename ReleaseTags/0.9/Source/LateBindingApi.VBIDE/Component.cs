//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class Component : _Component
	{
		#region Construction

		public Component(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Component(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Component(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Component()
		{
			CreateFromProgId("VBIDE.Component");
		}

		public Component(string progId)
		{
			CreateFromProgId(progId);
		}

		#endregion


	}
}
