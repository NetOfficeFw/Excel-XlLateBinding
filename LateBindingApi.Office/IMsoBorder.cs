//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class IMsoBorder : COMObject
	{
		#region Construction

		public IMsoBorder(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IMsoBorder(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IMsoBorder(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IMsoBorder()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public COMVariant Color
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Color");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Color", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant ColorIndex
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ColorIndex");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "ColorIndex", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant LineStyle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LineStyle");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "LineStyle", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Weight
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Weight");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Weight", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public COMObject Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF14")]
		public Int32 Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (Int32)returnValue;
			}
		}

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

		#endregion

		#region Methods

		#endregion

	}
}