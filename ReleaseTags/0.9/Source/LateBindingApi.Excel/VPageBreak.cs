//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class VPageBreak : COMObject
	{
		#region Construction

		public VPageBreak(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public VPageBreak(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public VPageBreak(COMObject replacedObject) : base(replacedObject)
		{
		}

		public VPageBreak()
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
		public LateBindingApi.Excel.Worksheet Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Worksheet newClass = new LateBindingApi.Excel.Worksheet(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlPageBreak Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Excel.Enums.XlPageBreak)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Type", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlPageBreakExtent Extent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Extent");
				return (LateBindingApi.Excel.Enums.XlPageBreakExtent)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range Location
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Location");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
			set
			{
				Invoker.PropertySet(this, "Location", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void DragOff(LateBindingApi.Excel.Enums.XlDirection direction, Int32 regionIndex)
		{
			object[] paramArray = new object[2];
			paramArray[0] = direction;
			paramArray[1] = regionIndex;
			Invoker.Method(this, "DragOff", paramArray);
		}

		#endregion

	}
}
