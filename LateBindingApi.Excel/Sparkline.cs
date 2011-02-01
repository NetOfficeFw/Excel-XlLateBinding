//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL14")]
	public class Sparkline : COMObject
	{
		#region Construction

		public Sparkline(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Sparkline(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Sparkline(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Sparkline()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL14")]
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

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Range Location
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Location");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
			set
			{
				Invoker.PropertySet(this, "Location", value);
			}						
		}


		[SupportByLibrary("XL14")]
		public string SourceData
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceData");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SourceData", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL14")]
		public void ModifyLocation(LateBindingApi.Excel.Range range)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(range,0);
			Invoker.Method(this, "ModifyLocation", paramArray);
		}

		[SupportByLibrary("XL14")]
		public void ModifySourceData(string formula)
		{
			object[] paramArray = new object[1];
			paramArray[0] = formula;
			Invoker.Method(this, "ModifySourceData", paramArray);
		}

		#endregion

	}
}
