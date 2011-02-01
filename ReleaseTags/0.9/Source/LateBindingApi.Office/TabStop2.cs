//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class TabStop2 : _IMsoDispObj
	{
		#region Construction

		public TabStop2(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public TabStop2(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public TabStop2(COMObject replacedObject) : base(replacedObject)
		{
		}

		public TabStop2()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Double Position
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Position");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Position", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoTabStopType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Office.Enums.MsoTabStopType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Type", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public void Clear()
		{
			Invoker.Method(this, "Clear", null);
		}

		#endregion

	}
}
