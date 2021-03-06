//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class MsoDebugOptions_UT : _IMsoDispObj
	{
		#region Construction

		public MsoDebugOptions_UT(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public MsoDebugOptions_UT(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public MsoDebugOptions_UT(COMObject replacedObject) : base(replacedObject)
		{
		}

		public MsoDebugOptions_UT()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string CollectionName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CollectionName");
				return (string)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.MsoDebugOptions_UTRunResult Run()
		{
			object returnValue = Invoker.MethodReturn(this, "Run", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.MsoDebugOptions_UTRunResult newClass = new LateBindingApi.Office.MsoDebugOptions_UTRunResult(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
