//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class _VBProject : _VBProject_Old
	{
		#region Construction

		public _VBProject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _VBProject(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _VBProject(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _VBProject()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Enums.Vbext_ProjectType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.VBIDE.Enums.Vbext_ProjectType)returnValue;
			}
		}

		[SupportByLibrary("VBE")]
		public string FileName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FileName");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("VBE")]
		public string BuildFileName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BuildFileName");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "BuildFileName", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("VBE")]
		public void SaveAs(string fileName)
		{
			object[] paramArray = new object[1];
			paramArray[0] = fileName;
			Invoker.Method(this, "SaveAs", paramArray);
		}

		[SupportByLibrary("VBE")]
		public void MakeCompiledFile()
		{
			Invoker.Method(this, "MakeCompiledFile", null);
		}

		#endregion

	}
}