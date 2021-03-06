//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class ReflectionFormat : _IMsoDispObj
	{
		#region Construction

		public ReflectionFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ReflectionFormat(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ReflectionFormat(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ReflectionFormat()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoReflectionType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Office.Enums.MsoReflectionType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Type", value);
			}						
		}


		[SupportByLibrary("OF14")]
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


		[SupportByLibrary("OF14")]
		public Double Size
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Size");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Size", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public Double Offset
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Offset");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Offset", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public Double Blur
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Blur");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Blur", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
