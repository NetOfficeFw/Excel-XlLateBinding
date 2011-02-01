//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL12","XL14")]
	public class ConditionValue : COMObject
	{
		#region Construction

		public ConditionValue(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ConditionValue(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ConditionValue(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ConditionValue()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL12","XL14")]
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

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlConditionValueTypes Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Excel.Enums.XlConditionValueTypes)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public COMVariant Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL12","XL14")]
		public void Modify(LateBindingApi.Excel.Enums.XlConditionValueTypes newtype)
		{
			object[] paramArray = new object[1];
			paramArray[0] = newtype;
			Invoker.Method(this, "Modify", paramArray);
		}

		[SupportByLibrary("XL12","XL14")]
		public void Modify(LateBindingApi.Excel.Enums.XlConditionValueTypes newtype, object newvalue)
		{
			object[] paramArray = new object[2];
			paramArray[0] = newtype;
			paramArray[1] = newvalue;
			Invoker.Method(this, "Modify", paramArray);
		}

		#endregion

	}
}