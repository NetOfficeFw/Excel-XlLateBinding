//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class WebComponentFormat : COMObject
	{
		#region Construction

		public WebComponentFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public WebComponentFormat(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public WebComponentFormat(COMObject replacedObject) : base(replacedObject)
		{
		}

		public WebComponentFormat()
		{
		}

		#endregion

		#region Properties

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
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string URL
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "URL");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "URL", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string HTML
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HTML");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HTML", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Name", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 Width
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Width");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Width", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 Height
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Height");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Height", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string PreviewGraphic
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PreviewGraphic");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PreviewGraphic", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void LaunchPropertiesWindow()
		{
			Invoker.Method(this, "LaunchPropertiesWindow", null);
		}

		#endregion

	}
}