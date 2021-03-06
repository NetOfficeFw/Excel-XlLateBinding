//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class ChartColorFormat : COMObject
	{
		#region Construction

		public ChartColorFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ChartColorFormat(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ChartColorFormat(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ChartColorFormat()
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 RGB
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RGB");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 _Default
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_Default");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (Int32)returnValue;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
