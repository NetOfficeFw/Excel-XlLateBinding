//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class WebComponentWindowExternal : COMObject
	{
		#region Construction

		public WebComponentWindowExternal(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public WebComponentWindowExternal(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public WebComponentWindowExternal(COMObject replacedObject) : base(replacedObject)
		{
		}

		public WebComponentWindowExternal()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 InterfaceVersion
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InterfaceVersion");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string ApplicationName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ApplicationName");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 ApplicationVersion
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ApplicationVersion");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.WebComponent WebComponent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebComponent");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.WebComponent newClass = new LateBindingApi.Office.WebComponent(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void CloseWindow()
		{
			Invoker.Method(this, "CloseWindow", null);
		}

		#endregion

	}
}
