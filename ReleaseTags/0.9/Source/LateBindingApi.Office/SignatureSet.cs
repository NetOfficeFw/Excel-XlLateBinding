//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class SignatureSet : _IMsoDispObj
	{
		#region Construction

		public SignatureSet(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SignatureSet(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SignatureSet(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SignatureSet()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.Signature returnClass = new LateBindingApi.Office.Signature (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.Signature this[Int32 iSig]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = iSig;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.Signature newClass = new LateBindingApi.Office.Signature(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
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
		public bool CanAddSignatureLine
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CanAddSignatureLine");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoSignatureSubset Subset
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Subset");
				return (LateBindingApi.Office.Enums.MsoSignatureSubset)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Subset", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public bool ShowSignaturesPane
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShowSignaturesPane");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ShowSignaturesPane", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public LateBindingApi.Office.Signature Add()
		{
			object returnValue = Invoker.MethodReturn(this, "Add", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.Signature newClass = new LateBindingApi.Office.Signature(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void Commit()
		{
			Invoker.Method(this, "Commit", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Signature AddNonVisibleSignature()
		{
			object returnValue = Invoker.MethodReturn(this, "AddNonVisibleSignature", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.Signature newClass = new LateBindingApi.Office.Signature(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Signature AddNonVisibleSignature(object varSigProv)
		{
			object[] paramArray = new object[1];
			paramArray[0] = varSigProv;
			object returnValue = Invoker.MethodReturn(this, "AddNonVisibleSignature", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.Signature newClass = new LateBindingApi.Office.Signature(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Signature AddSignatureLine()
		{
			object returnValue = Invoker.MethodReturn(this, "AddSignatureLine", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.Signature newClass = new LateBindingApi.Office.Signature(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Signature AddSignatureLine(object varSigProv)
		{
			object[] paramArray = new object[1];
			paramArray[0] = varSigProv;
			object returnValue = Invoker.MethodReturn(this, "AddSignatureLine", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.Signature newClass = new LateBindingApi.Office.Signature(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
