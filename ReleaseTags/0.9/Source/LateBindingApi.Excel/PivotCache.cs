//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class PivotCache : COMObject
	{
		#region Construction

		public PivotCache(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PivotCache(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PivotCache(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PivotCache()
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
		public bool BackgroundQuery
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BackgroundQuery");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "BackgroundQuery", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Connection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Connection");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Connection", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool EnableRefresh
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "EnableRefresh");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "EnableRefresh", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Index
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Index");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 MemoryUsed
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MemoryUsed");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool OptimizeCache
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OptimizeCache");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OptimizeCache", value);
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
		public COMVariant RefreshDate
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RefreshDate");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string RefreshName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RefreshName");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool RefreshOnFileOpen
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RefreshOnFileOpen");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RefreshOnFileOpen", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Sql
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Sql");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Sql", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool SavePassword
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SavePassword");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SavePassword", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SourceData
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceData");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "SourceData", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant CommandText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CommandText");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "CommandText", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlCmdType CommandType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CommandType");
				return (LateBindingApi.Excel.Enums.XlCmdType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CommandType", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlQueryType QueryType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "QueryType");
				return (LateBindingApi.Excel.Enums.XlQueryType)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool MaintainConnection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MaintainConnection");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "MaintainConnection", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 RefreshPeriod
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RefreshPeriod");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RefreshPeriod", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Recordset
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Recordset");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Recordset", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant LocalConnection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LocalConnection");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "LocalConnection", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool UseLocalConnection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "UseLocalConnection");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "UseLocalConnection", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMObject ADOConnection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ADOConnection");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public bool IsConnected
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsConnected");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public bool OLAP
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OLAP");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlPivotTableSourceType SourceType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceType");
				return (LateBindingApi.Excel.Enums.XlPivotTableSourceType)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlPivotTableMissingItems MissingItemsLimit
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MissingItemsLimit");
				return (LateBindingApi.Excel.Enums.XlPivotTableMissingItems)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "MissingItemsLimit", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string SourceConnectionFile
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceConnectionFile");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SourceConnectionFile", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string SourceDataFile
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SourceDataFile");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlRobustConnect RobustConnect
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RobustConnect");
				return (LateBindingApi.Excel.Enums.XlRobustConnect)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RobustConnect", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.WorkbookConnection WorkbookConnection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WorkbookConnection");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.WorkbookConnection newClass = new LateBindingApi.Excel.WorkbookConnection(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlPivotTableVersionList Version
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Version");
				return (LateBindingApi.Excel.Enums.XlPivotTableVersionList)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public bool UpgradeOnRefresh
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "UpgradeOnRefresh");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "UpgradeOnRefresh", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Refresh()
		{
			Invoker.Method(this, "Refresh", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void ResetTimer()
		{
			Invoker.Method(this, "ResetTimer", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.PivotTable CreatePivotTable(object tableDestination)
		{
			object[] paramArray = new object[1];
			paramArray[0] = tableDestination;
			object returnValue = Invoker.MethodReturn(this, "CreatePivotTable", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.PivotTable newClass = new LateBindingApi.Excel.PivotTable(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL9")]
		public LateBindingApi.Excel.PivotTable CreatePivotTable(object tableDestination, object tableName, object readData)
		{
			object[] paramArray = new object[3];
			paramArray[0] = tableDestination;
			paramArray[1] = tableName;
			paramArray[2] = readData;
			object returnValue = Invoker.MethodReturn(this, "CreatePivotTable", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.PivotTable newClass = new LateBindingApi.Excel.PivotTable(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.PivotTable CreatePivotTable(object tableDestination, object tableName, object readData, object defaultVersion)
		{
			object[] paramArray = new object[4];
			paramArray[0] = tableDestination;
			paramArray[1] = tableName;
			paramArray[2] = readData;
			paramArray[3] = defaultVersion;
			object returnValue = Invoker.MethodReturn(this, "CreatePivotTable", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.PivotTable newClass = new LateBindingApi.Excel.PivotTable(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void MakeConnection()
		{
			Invoker.Method(this, "MakeConnection", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void SaveAsODC(string oDCFileName)
		{
			object[] paramArray = new object[1];
			paramArray[0] = oDCFileName;
			Invoker.Method(this, "SaveAsODC", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public void SaveAsODC(string oDCFileName, object description, object keywords)
		{
			object[] paramArray = new object[3];
			paramArray[0] = oDCFileName;
			paramArray[1] = description;
			paramArray[2] = keywords;
			Invoker.Method(this, "SaveAsODC", paramArray);
		}

		#endregion

	}
}
