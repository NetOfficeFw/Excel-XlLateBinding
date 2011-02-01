//Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class Global : _Global
	{
		#region Construction

		public Global(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Global(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Global(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Global()
		{
			CreateFromProgId("Excel.Global");
		}

		public Global(string progId)
		{
			CreateFromProgId(progId);
		}

		#endregion


	}
}
