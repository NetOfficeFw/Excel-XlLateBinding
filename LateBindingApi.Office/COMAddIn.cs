//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class COMAddIn : _IMsoDispObj
	{
		#region Construction

		public COMAddIn(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public COMAddIn(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public COMAddIn(COMObject replacedObject) : base(replacedObject)
		{
		}

		public COMAddIn()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Description
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Description");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Description", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string ProgId
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ProgId");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Guid
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Guid");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool Connect
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Connect");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Connect", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMObject Object
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Object");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Object", value);
			}						
		}


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

		#endregion

		#region Methods

		#endregion

	}
}
