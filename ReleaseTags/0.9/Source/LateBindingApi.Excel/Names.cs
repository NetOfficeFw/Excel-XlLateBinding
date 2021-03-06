//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class Names : COMObject
	{
		#region Construction

		public Names(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Names(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Names(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Names()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Excel.Name returnClass = new LateBindingApi.Excel.Name (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Name Add()
		{
			object returnValue = Invoker.MethodReturn(this, "Add", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Name newClass = new LateBindingApi.Excel.Name(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local)
		{
			object[] paramArray = new object[11];
			paramArray[0] = name;
			paramArray[1] = refersTo;
			paramArray[2] = visible;
			paramArray[3] = macroType;
			paramArray[4] = shortcutKey;
			paramArray[5] = category;
			paramArray[6] = nameLocal;
			paramArray[7] = refersToLocal;
			paramArray[8] = categoryLocal;
			paramArray[9] = refersToR1C1;
			paramArray[10] = refersToR1C1Local;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Name newClass = new LateBindingApi.Excel.Name(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Name this[object index, object indexLocal, object refersTo]
		{
			get
			{
				object[] paramArray = new object[3];
				paramArray[0] = index;
				paramArray[1] = indexLocal;
				paramArray[2] = refersTo;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Name newClass = new LateBindingApi.Excel.Name(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Name _Default()
		{
			object returnValue = Invoker.MethodReturn(this, "_Default", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Name newClass = new LateBindingApi.Excel.Name(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Name _Default(object index, object indexLocal, object refersTo)
		{
			object[] paramArray = new object[3];
			paramArray[0] = index;
			paramArray[1] = indexLocal;
			paramArray[2] = refersTo;
			object returnValue = Invoker.MethodReturn(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Name newClass = new LateBindingApi.Excel.Name(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
