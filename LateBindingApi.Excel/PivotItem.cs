//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class PivotItem : COMObject
	{
		#region Construction

		public PivotItem(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PivotItem(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PivotItem(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PivotItem()
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
		public LateBindingApi.Excel.PivotField Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.PivotField newClass = new LateBindingApi.Excel.PivotField(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ChildItems
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ChildItems", null);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant get_ChildItems(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.PropertyGet(this, "ChildItems", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range DataRange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DataRange");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string _Default
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_Default");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "_Default", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range LabelRange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LabelRange");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Name", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.PivotItem ParentItem
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ParentItem");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.PivotItem newClass = new LateBindingApi.Excel.PivotItem(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool ParentShowDetail
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ParentShowDetail");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Position
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Position");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Position", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool ShowDetail
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShowDetail");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ShowDetail", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SourceName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceName");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Value
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Value");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Value", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Visible
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Visible");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Visible", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool IsCalculated
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsCalculated");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 RecordCount
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RecordCount");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Formula
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Formula");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Formula", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Caption
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Caption");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Caption", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DrilledDown
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DrilledDown");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DrilledDown", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string StandardFormula
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "StandardFormula");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "StandardFormula", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string SourceNameStandard
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceNameStandard");
				return (string)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("XL12","XL14")]
		public void DrillTo(string field)
		{
			object[] paramArray = new object[1];
			paramArray[0] = field;
			Invoker.Method(this, "DrillTo", paramArray);
		}

		#endregion

	}
}
