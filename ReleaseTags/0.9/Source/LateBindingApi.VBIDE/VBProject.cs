//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class VBProject : _VBProject
	{
		#region Construction

		public VBProject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public VBProject(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public VBProject(COMObject replacedObject) : base(replacedObject)
		{
		}

		public VBProject()
		{
			CreateFromProgId("VBIDE.VBProject");
		}

		public VBProject(string progId)
		{
			CreateFromProgId(progId);
		}

		#endregion


	}
}
