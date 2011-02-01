//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class MsoDebugOptions_UTManager : _IMsoDispObj
	{
		#region Construction

		public MsoDebugOptions_UTManager(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public MsoDebugOptions_UTManager(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public MsoDebugOptions_UTManager(COMObject replacedObject) : base(replacedObject)
		{
		}

		public MsoDebugOptions_UTManager()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.MsoDebugOptions_UTs UnitTests
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "UnitTests");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.MsoDebugOptions_UTs newClass = new LateBindingApi.Office.MsoDebugOptions_UTs(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public bool ReportErrors
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ReportErrors");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ReportErrors", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public void NotifyStartOfTestSuiteRun()
		{
			Invoker.Method(this, "NotifyStartOfTestSuiteRun", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void NotifyEndOfTestSuiteRun()
		{
			Invoker.Method(this, "NotifyEndOfTestSuiteRun", null);
		}

		#endregion

	}
}
