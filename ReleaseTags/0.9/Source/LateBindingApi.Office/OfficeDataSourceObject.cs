//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class OfficeDataSourceObject : COMObject
	{
		#region Construction

		public OfficeDataSourceObject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public OfficeDataSourceObject(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public OfficeDataSourceObject(COMObject replacedObject) : base(replacedObject)
		{
		}

		public OfficeDataSourceObject()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string ConnectString
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ConnectString");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ConnectString", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string Table
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Table");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Table", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string DataSource
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DataSource");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DataSource", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject Columns
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Columns");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 RowCount
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RowCount");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject Filters
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Filters");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 Move(LateBindingApi.Office.Enums.MsoMoveRow msoMoveRow, Int32 rowNbr)
		{
			object[] paramArray = new object[2];
			paramArray[0] = msoMoveRow;
			paramArray[1] = rowNbr;
			object returnValue = Invoker.MethodReturn(this, "Move", paramArray);
			return (Int32)returnValue;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void Open(string bstrSrc, string bstrConnect, string bstrTable, Int32 fOpenExclusive, Int32 fNeverPrompt)
		{
			object[] paramArray = new object[5];
			paramArray[0] = bstrSrc;
			paramArray[1] = bstrConnect;
			paramArray[2] = bstrTable;
			paramArray[3] = fOpenExclusive;
			paramArray[4] = fNeverPrompt;
			Invoker.Method(this, "Open", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void SetSortOrder(string sortField1, bool sortAscending1, string sortField2, bool sortAscending2, string sortField3, bool sortAscending3)
		{
			object[] paramArray = new object[6];
			paramArray[0] = sortField1;
			paramArray[1] = sortAscending1;
			paramArray[2] = sortField2;
			paramArray[3] = sortAscending2;
			paramArray[4] = sortField3;
			paramArray[5] = sortAscending3;
			Invoker.Method(this, "SetSortOrder", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void ApplyFilter()
		{
			Invoker.Method(this, "ApplyFilter", null);
		}

		#endregion

	}
}
