//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF14")]
	public class PickerResult : _IMsoDispObj
	{
		#region Construction

		public PickerResult(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PickerResult(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PickerResult(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PickerResult()
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
		public string DisplayName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayName");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayName", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public string Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Type", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public string SIPId
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SIPId");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SIPId", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public COMVariant ItemData
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ItemData");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "ItemData", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public COMVariant SubItems
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SubItems");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "SubItems", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public COMVariant DuplicateResults
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DuplicateResults");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF14")]
		public LateBindingApi.Office.PickerFields Fields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Fields");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.PickerFields newClass = new LateBindingApi.Office.PickerFields(this, returnValue);
				return newClass;
			}
			set
			{
				Invoker.PropertySet(this, "Fields", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
