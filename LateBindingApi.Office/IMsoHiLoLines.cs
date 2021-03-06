//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class IMsoHiLoLines : COMObject
	{
		#region Construction

		public IMsoHiLoLines(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IMsoHiLoLines(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IMsoHiLoLines(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IMsoHiLoLines()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.IMsoBorder Border
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Border");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IMsoBorder newClass = new LateBindingApi.Office.IMsoBorder(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.IMsoChartFormat Format
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Format");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IMsoChartFormat newClass = new LateBindingApi.Office.IMsoChartFormat(this, returnValue);
				return newClass;
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

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public void Select()
		{
			Invoker.Method(this, "Select", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		#endregion

	}
}
