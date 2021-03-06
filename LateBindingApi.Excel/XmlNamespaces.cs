//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL11","XL12","XL14")]
	public class XmlNamespaces : COMObject
	{
		#region Construction

		public XmlNamespaces(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public XmlNamespaces(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public XmlNamespaces(COMObject replacedObject) : base(replacedObject)
		{
		}

		public XmlNamespaces()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL11","XL12","XL14")]
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

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.XmlNamespace get__Default(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.XmlNamespace newClass = new LateBindingApi.Excel.XmlNamespace(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.XmlNamespace this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.XmlNamespace newClass = new LateBindingApi.Excel.XmlNamespace(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public string Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Excel.XmlNamespace returnClass = new LateBindingApi.Excel.XmlNamespace (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL11","XL12","XL14")]
		public void InstallManifest(string path)
		{
			object[] paramArray = new object[1];
			paramArray[0] = path;
			Invoker.Method(this, "InstallManifest", paramArray);
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public void InstallManifest(string path, object installForAllUsers)
		{
			object[] paramArray = new object[2];
			paramArray[0] = path;
			paramArray[1] = installForAllUsers;
			Invoker.Method(this, "InstallManifest", paramArray);
		}

		#endregion

	}
}
