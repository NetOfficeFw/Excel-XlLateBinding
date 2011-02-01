//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class ColorFormat : _IMsoDispObj
	{
		#region Construction

		public ColorFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ColorFormat(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ColorFormat(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ColorFormat()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 RGB
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RGB");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RGB", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 SchemeColor
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SchemeColor");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SchemeColor", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoColorType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Office.Enums.MsoColorType)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Double TintAndShade
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TintAndShade");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TintAndShade", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoThemeColorIndex ObjectThemeColor
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ObjectThemeColor");
				return (LateBindingApi.Office.Enums.MsoThemeColorIndex)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ObjectThemeColor", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public Double Brightness
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Brightness");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Brightness", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
