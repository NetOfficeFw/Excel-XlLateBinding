//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF14")]
	public class PickerField : _IMsoDispObj
	{
		#region Construction

		public PickerField(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PickerField(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PickerField(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PickerField()
		{
		}

		#endregion

		#region Properties

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
		public LateBindingApi.Office.Enums.MsoPickerField Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Office.Enums.MsoPickerField)returnValue;
			}
		}

		[SupportByLibrary("OF14")]
		public bool IsHidden
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsHidden");
				return (bool)returnValue;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
