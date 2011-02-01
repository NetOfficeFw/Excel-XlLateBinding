//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF14")]
	public class PickerProperty : _IMsoDispObj
	{
		#region Construction

		public PickerProperty(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PickerProperty(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PickerProperty(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PickerProperty()
		{
		}

		#endregion

		#region Properties

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
		public COMVariant Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
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

		#endregion

		#region Methods

		#endregion

	}
}
