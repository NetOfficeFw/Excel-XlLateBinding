//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class VBProjects : _VBProjects
	{
		#region Construction

		public VBProjects(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public VBProjects(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public VBProjects(COMObject replacedObject) : base(replacedObject)
		{
		}

		public VBProjects()
		{
			CreateFromProgId("VBIDE.VBProjects");
		}

		public VBProjects(string progId)
		{
			CreateFromProgId(progId);
		}

		#endregion


	}
}
