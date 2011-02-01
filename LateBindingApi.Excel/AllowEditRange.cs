//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14")]
	public class AllowEditRange : COMObject
	{
		#region Construction

		public AllowEditRange(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public AllowEditRange(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public AllowEditRange(COMObject replacedObject) : base(replacedObject)
		{
		}

		public AllowEditRange()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string Title
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Title");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Title", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Range Range
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Range");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
			set
			{
				Invoker.PropertySet(this, "Range", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.UserAccessList Users
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Users");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.UserAccessList newClass = new LateBindingApi.Excel.UserAccessList(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void ChangePassword(string password)
		{
			object[] paramArray = new object[1];
			paramArray[0] = password;
			Invoker.Method(this, "ChangePassword", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void Unprotect()
		{
			Invoker.Method(this, "Unprotect", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void Unprotect(object password)
		{
			object[] paramArray = new object[1];
			paramArray[0] = password;
			Invoker.Method(this, "Unprotect", paramArray);
		}

		#endregion

	}
}