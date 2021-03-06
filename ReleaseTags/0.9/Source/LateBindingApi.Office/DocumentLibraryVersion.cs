//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF11","OF12","OF14")]
	public class DocumentLibraryVersion : _IMsoDispObj
	{
		#region Construction

		public DocumentLibraryVersion(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public DocumentLibraryVersion(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public DocumentLibraryVersion(COMObject replacedObject) : base(replacedObject)
		{
		}

		public DocumentLibraryVersion()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMVariant Modified
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Modified");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public Int32 Index
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Index");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public string ModifiedBy
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ModifiedBy");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public string Comments
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Comments");
				return (string)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF11","OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMObject Open()
		{
			object returnValue = Invoker.MethodReturn(this, "Open", null);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMObject Restore()
		{
			object returnValue = Invoker.MethodReturn(this, "Restore", null);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
