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
	public class ListObjects : COMObject
	{
		#region Construction

		public ListObjects(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ListObjects(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ListObjects(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ListObjects()
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
		public LateBindingApi.Excel.ListObject get__Default(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
			return newClass;
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
				LateBindingApi.Excel.ListObject returnClass = new LateBindingApi.Excel.ListObject (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.ListObject this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
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

		#endregion

		#region Methods

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.ListObject Add(LateBindingApi.Excel.Enums.XlListObjectSourceType sourceType, object source, object linkSource, LateBindingApi.Excel.Enums.XlYesNoGuess xlListObjectHasHeaders)
		{
			object[] paramArray = new object[4];
			paramArray[0] = sourceType;
			paramArray[1] = source;
			paramArray[2] = linkSource;
			paramArray[3] = xlListObjectHasHeaders;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL11")]
		public LateBindingApi.Excel.ListObject Add(LateBindingApi.Excel.Enums.XlListObjectSourceType sourceType, object source, object linkSource, LateBindingApi.Excel.Enums.XlYesNoGuess xlListObjectHasHeaders, object destination)
		{
			object[] paramArray = new object[5];
			paramArray[0] = sourceType;
			paramArray[1] = source;
			paramArray[2] = linkSource;
			paramArray[3] = xlListObjectHasHeaders;
			paramArray[4] = destination;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.ListObject Add(LateBindingApi.Excel.Enums.XlListObjectSourceType sourceType, object source, object linkSource, LateBindingApi.Excel.Enums.XlYesNoGuess xlListObjectHasHeaders, object destination, object tableStyleName)
		{
			object[] paramArray = new object[6];
			paramArray[0] = sourceType;
			paramArray[1] = source;
			paramArray[2] = linkSource;
			paramArray[3] = xlListObjectHasHeaders;
			paramArray[4] = destination;
			paramArray[5] = tableStyleName;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.ListObject _Add(LateBindingApi.Excel.Enums.XlListObjectSourceType sourceType, object source, object linkSource, LateBindingApi.Excel.Enums.XlYesNoGuess xlListObjectHasHeaders)
		{
			object[] paramArray = new object[4];
			paramArray[0] = sourceType;
			paramArray[1] = source;
			paramArray[2] = linkSource;
			paramArray[3] = xlListObjectHasHeaders;
			object returnValue = Invoker.MethodReturn(this, "_Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.ListObject _Add(LateBindingApi.Excel.Enums.XlListObjectSourceType sourceType, object source, object linkSource, LateBindingApi.Excel.Enums.XlYesNoGuess xlListObjectHasHeaders, object destination)
		{
			object[] paramArray = new object[5];
			paramArray[0] = sourceType;
			paramArray[1] = source;
			paramArray[2] = linkSource;
			paramArray[3] = xlListObjectHasHeaders;
			paramArray[4] = destination;
			object returnValue = Invoker.MethodReturn(this, "_Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
			return newClass;
		}

		#endregion

	}
}