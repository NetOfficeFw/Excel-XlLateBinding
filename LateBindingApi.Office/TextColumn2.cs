//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class TextColumn2 : _IMsoDispObj
	{
		#region Construction

		public TextColumn2(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public TextColumn2(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public TextColumn2(COMObject replacedObject) : base(replacedObject)
		{
		}

		public TextColumn2()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public Int32 Number
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Number");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Number", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public Double Spacing
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Spacing");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Spacing", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoTextDirection TextDirection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextDirection");
				return (LateBindingApi.Office.Enums.MsoTextDirection)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextDirection", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
