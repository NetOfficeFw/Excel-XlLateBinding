//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class _CustomXMLParts : _IMsoDispObj
	{
		#region Construction

		public _CustomXMLParts(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _CustomXMLParts(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _CustomXMLParts(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _CustomXMLParts()
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
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLPart this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CustomXMLPart newClass = new LateBindingApi.Office.CustomXMLPart(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.CustomXMLPart returnClass = new LateBindingApi.Office.CustomXMLPart (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLPart Add(string xML)
		{
			object[] paramArray = new object[1];
			paramArray[0] = xML;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CustomXMLPart newClass = new LateBindingApi.Office.CustomXMLPart(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLPart Add(string xML, object schemaCollection)
		{
			object[] paramArray = new object[2];
			paramArray[0] = xML;
			paramArray[1] = schemaCollection;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CustomXMLPart newClass = new LateBindingApi.Office.CustomXMLPart(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLPart SelectByID(string id)
		{
			object[] paramArray = new object[1];
			paramArray[0] = id;
			object returnValue = Invoker.MethodReturn(this, "SelectByID", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CustomXMLPart newClass = new LateBindingApi.Office.CustomXMLPart(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.CustomXMLParts SelectByNamespace(string namespaceURI)
		{
			object[] paramArray = new object[1];
			paramArray[0] = namespaceURI;
			object returnValue = Invoker.MethodReturn(this, "SelectByNamespace", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CustomXMLParts newClass = new LateBindingApi.Office.CustomXMLParts(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
