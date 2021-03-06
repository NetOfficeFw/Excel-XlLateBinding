//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class ILicValidator : COMObject
	{
		#region Construction

		public ILicValidator(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ILicValidator(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ILicValidator(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ILicValidator()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMVariant Products
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Products");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 Selection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Selection");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Selection", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
