//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL14")]
	public class Slicers : COMObject
	{
		#region Construction

		public Slicers(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Slicers(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Slicers(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Slicers()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL14")]
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

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL14")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Slicer this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Slicer newClass = new LateBindingApi.Excel.Slicer(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Slicer get__Default(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "_Default", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Slicer newClass = new LateBindingApi.Excel.Slicer(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Excel.Slicer returnClass = new LateBindingApi.Excel.Slicer (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Slicer Add(object slicerDestination)
		{
			object[] paramArray = new object[1];
			paramArray[0] = slicerDestination;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Slicer newClass = new LateBindingApi.Excel.Slicer(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL14")]
		public LateBindingApi.Excel.Slicer Add(object slicerDestination, object level, object name, object caption, object top, object left, object width, object height)
		{
			object[] paramArray = new object[8];
			paramArray[0] = slicerDestination;
			paramArray[1] = level;
			paramArray[2] = name;
			paramArray[3] = caption;
			paramArray[4] = top;
			paramArray[5] = left;
			paramArray[6] = width;
			paramArray[7] = height;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Slicer newClass = new LateBindingApi.Excel.Slicer(this, returnValue);
			return newClass;
		}

		#endregion

	}
}
