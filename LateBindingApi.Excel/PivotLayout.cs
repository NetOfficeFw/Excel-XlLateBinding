//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class PivotLayout : COMObject
	{
		#region Construction

		public PivotLayout(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PivotLayout(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PivotLayout(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PivotLayout()
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
		public COMObject ColumnFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ColumnFields", null);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject get_ColumnFields(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "ColumnFields", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject DataFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DataFields", null);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject get_DataFields(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "DataFields", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject PageFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PageFields", null);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject get_PageFields(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "PageFields", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject RowFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RowFields", null);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject get_RowFields(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "RowFields", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject HiddenFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HiddenFields", null);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject get_HiddenFields(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "HiddenFields", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject VisibleFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VisibleFields", null);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject get_VisibleFields(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "VisibleFields", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject PivotFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PivotFields", null);
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject get_PivotFields(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "PivotFields", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.CubeFields CubeFields
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CubeFields");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.CubeFields newClass = new LateBindingApi.Excel.CubeFields(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.PivotCache PivotCache
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PivotCache");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.PivotCache newClass = new LateBindingApi.Excel.PivotCache(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.PivotTable PivotTable
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PivotTable");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.PivotTable newClass = new LateBindingApi.Excel.PivotTable(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string InnerDetail
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InnerDetail");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InnerDetail", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void AddFields()
		{
			Invoker.Method(this, "AddFields", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void AddFields(object rowFields, object columnFields, object pageFields, object appendField)
		{
			object[] paramArray = new object[4];
			paramArray[0] = rowFields;
			paramArray[1] = columnFields;
			paramArray[2] = pageFields;
			paramArray[3] = appendField;
			Invoker.Method(this, "AddFields", paramArray);
		}

		#endregion

	}
}
