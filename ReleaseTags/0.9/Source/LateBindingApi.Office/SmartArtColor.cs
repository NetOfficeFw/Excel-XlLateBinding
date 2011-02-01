//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF14")]
	public class SmartArtColor : _IMsoDispObj
	{
		#region Construction

		public SmartArtColor(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SmartArtColor(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SmartArtColor(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SmartArtColor()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF14")]
		public string Id
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Id");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF14")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF14")]
		public string Description
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Description");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF14")]
		public string Category
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Category");
				return (string)returnValue;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
