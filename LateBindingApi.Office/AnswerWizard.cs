//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class AnswerWizard : _IMsoDispObj
	{
		#region Construction

		public AnswerWizard(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public AnswerWizard(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public AnswerWizard(COMObject replacedObject) : base(replacedObject)
		{
		}

		public AnswerWizard()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.AnswerWizardFiles Files
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Files");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.AnswerWizardFiles newClass = new LateBindingApi.Office.AnswerWizardFiles(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void ClearFileList()
		{
			Invoker.Method(this, "ClearFileList", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void ResetFileList()
		{
			Invoker.Method(this, "ResetFileList", null);
		}

		#endregion

	}
}
