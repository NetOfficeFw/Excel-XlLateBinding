//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14")]
	public class AutoRecover : COMObject
	{
		#region Construction

		public AutoRecover(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public AutoRecover(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public AutoRecover(COMObject replacedObject) : base(replacedObject)
		{
		}

		public AutoRecover()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
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

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public bool Enabled
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Enabled");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Enabled", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Int32 Time
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Time");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Time", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string Path
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Path");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Path", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
