//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class GradientStop : _IMsoDispObj
	{
		#region Construction

		public GradientStop(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public GradientStop(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public GradientStop(COMObject replacedObject) : base(replacedObject)
		{
		}

		public GradientStop()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.ColorFormat Color
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Color");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.ColorFormat newClass = new LateBindingApi.Office.ColorFormat(this, returnValue);
				return newClass;
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
		public Double Transparency
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Transparency");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Transparency", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
