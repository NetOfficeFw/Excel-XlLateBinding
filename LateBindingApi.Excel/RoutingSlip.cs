//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class RoutingSlip : COMObject
	{
		#region Construction

		public RoutingSlip(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public RoutingSlip(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public RoutingSlip(COMObject replacedObject) : base(replacedObject)
		{
		}

		public RoutingSlip()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Application Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Application newClass = new LateBindingApi.Excel.Application(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlRoutingSlipDelivery Delivery
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Delivery");
				return (LateBindingApi.Excel.Enums.XlRoutingSlipDelivery)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Delivery", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Message
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Message");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Message", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Recipients
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Recipients", null);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Recipients", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant get_Recipients(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
				object returnValue = Invoker.PropertyGet(this, "Recipients", paramArray);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
		}
		
		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void set_Recipients(object index, object value)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.PropertySet(this, "Recipients", paramArray, value);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool ReturnWhenDone
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ReturnWhenDone");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ReturnWhenDone", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlRoutingSlipStatus Status
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Status");
				return (LateBindingApi.Excel.Enums.XlRoutingSlipStatus)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Subject
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Subject");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Subject", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool TrackStatus
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TrackStatus");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TrackStatus", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Reset()
		{
			object returnValue = Invoker.MethodReturn(this, "Reset", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
