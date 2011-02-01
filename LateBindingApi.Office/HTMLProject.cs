//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class HTMLProject : _IMsoDispObj
	{
		#region Construction

		public HTMLProject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public HTMLProject(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public HTMLProject(COMObject replacedObject) : base(replacedObject)
		{
		}

		public HTMLProject()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoHTMLProjectState State
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "State");
				return (LateBindingApi.Office.Enums.MsoHTMLProjectState)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.HTMLProjectItems HTMLProjectItems
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HTMLProjectItems");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.HTMLProjectItems newClass = new LateBindingApi.Office.HTMLProjectItems(this, returnValue);
				return newClass;
			}
		}

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

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void RefreshProject(bool refresh)
		{
			object[] paramArray = new object[1];
			paramArray[0] = refresh;
			Invoker.Method(this, "RefreshProject", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void RefreshDocument(bool refresh)
		{
			object[] paramArray = new object[1];
			paramArray[0] = refresh;
			Invoker.Method(this, "RefreshDocument", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Open(LateBindingApi.Office.Enums.MsoHTMLProjectOpen openKind)
		{
			object[] paramArray = new object[1];
			paramArray[0] = openKind;
			Invoker.Method(this, "Open", paramArray);
		}

		#endregion

	}
}
