//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class Adjustments : _IMsoDispObj
	{
		#region Construction

		public Adjustments(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Adjustments(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Adjustments(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Adjustments()
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
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Double get_Item(Int32 index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				return (Double)returnValue;
		}
		
		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void set_Item(Int32 index, object value)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.PropertySet(this, "Item", paramArray, value);
		}

		#endregion

		#region Methods

		#endregion

	}
}