//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class IMsoEnvelopeVB : COMObject
	{
		#region Construction

		public IMsoEnvelopeVB(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IMsoEnvelopeVB(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IMsoEnvelopeVB(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IMsoEnvelopeVB()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string Introduction
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Introduction");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Introduction", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject Item
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Item");
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
		public COMObject CommandBars
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CommandBars");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
