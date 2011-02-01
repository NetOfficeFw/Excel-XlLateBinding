//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class _QueryTable : COMObject
	{
		#region Construction

		public _QueryTable(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _QueryTable(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _QueryTable(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _QueryTable()
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
		public bool FieldNames
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FieldNames");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FieldNames", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool RowNumbers
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RowNumbers");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RowNumbers", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool FillAdjacentFormulas
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FillAdjacentFormulas");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FillAdjacentFormulas", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool HasAutoFormat
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasAutoFormat");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasAutoFormat", value);
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
		public bool Refreshing
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Refreshing");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool FetchedRowOverflow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FetchedRowOverflow");
				return (bool)returnValue;
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
		public LateBindingApi.Excel.Enums.XlCellInsertionMode RefreshStyle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RefreshStyle");
				return (LateBindingApi.Excel.Enums.XlCellInsertionMode)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RefreshStyle", value);
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
		public LateBindingApi.Excel.Range Destination
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Destination");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
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
		public string PostText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PostText");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PostText", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range ResultRange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ResultRange");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Parameters Parameters
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parameters");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Parameters newClass = new LateBindingApi.Excel.Parameters(this, returnValue);
				return newClass;
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
		public bool SaveData
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SaveData");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SaveData", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool TablesOnlyFromHTML
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TablesOnlyFromHTML");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TablesOnlyFromHTML", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool EnableEditing
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "EnableEditing");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "EnableEditing", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 TextFilePlatform
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFilePlatform");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFilePlatform", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 TextFileStartRow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileStartRow");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileStartRow", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlTextParsingType TextFileParseType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileParseType");
				return (LateBindingApi.Excel.Enums.XlTextParsingType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileParseType", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlTextQualifier TextFileTextQualifier
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileTextQualifier");
				return (LateBindingApi.Excel.Enums.XlTextQualifier)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileTextQualifier", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool TextFileConsecutiveDelimiter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileConsecutiveDelimiter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileConsecutiveDelimiter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool TextFileTabDelimiter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileTabDelimiter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileTabDelimiter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool TextFileSemicolonDelimiter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileSemicolonDelimiter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileSemicolonDelimiter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool TextFileCommaDelimiter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileCommaDelimiter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileCommaDelimiter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool TextFileSpaceDelimiter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileSpaceDelimiter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileSpaceDelimiter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string TextFileOtherDelimiter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileOtherDelimiter");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileOtherDelimiter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant TextFileColumnDataTypes
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileColumnDataTypes");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileColumnDataTypes", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant TextFileFixedColumnWidths
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileFixedColumnWidths");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileFixedColumnWidths", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool PreserveColumnInfo
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PreserveColumnInfo");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PreserveColumnInfo", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool PreserveFormatting
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PreserveFormatting");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PreserveFormatting", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool AdjustColumnWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AdjustColumnWidth");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AdjustColumnWidth", value);
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
		public bool TextFilePromptOnRefresh
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFilePromptOnRefresh");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFilePromptOnRefresh", value);
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
		public string TextFileDecimalSeparator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileDecimalSeparator");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileDecimalSeparator", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string TextFileThousandsSeparator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileThousandsSeparator");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileThousandsSeparator", value);
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
		public LateBindingApi.Excel.Enums.XlWebSelectionType WebSelectionType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebSelectionType");
				return (LateBindingApi.Excel.Enums.XlWebSelectionType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebSelectionType", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlWebFormatting WebFormatting
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebFormatting");
				return (LateBindingApi.Excel.Enums.XlWebFormatting)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebFormatting", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string WebTables
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebTables");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebTables", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool WebPreFormattedTextToColumns
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebPreFormattedTextToColumns");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebPreFormattedTextToColumns", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool WebSingleBlockTextImport
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebSingleBlockTextImport");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebSingleBlockTextImport", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool WebDisableDateRecognition
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebDisableDateRecognition");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebDisableDateRecognition", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool WebConsecutiveDelimitersAsOne
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebConsecutiveDelimitersAsOne");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebConsecutiveDelimitersAsOne", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public bool WebDisableRedirections
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WebDisableRedirections");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WebDisableRedirections", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMVariant EditWebPage
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "EditWebPage");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "EditWebPage", value);
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
			set
			{
				Invoker.PropertySet(this, "SourceDataFile", value);
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


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public bool TextFileTrailingMinusNumbers
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileTrailingMinusNumbers");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileTrailingMinusNumbers", value);
			}						
		}


		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.ListObject ListObject
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListObject");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.ListObject newClass = new LateBindingApi.Excel.ListObject(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlTextVisualLayoutType TextFileVisualLayout
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TextFileVisualLayout");
				return (LateBindingApi.Excel.Enums.XlTextVisualLayoutType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TextFileVisualLayout", value);
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
		public LateBindingApi.Excel.Sort Sort
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Sort");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Sort newClass = new LateBindingApi.Excel.Sort(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void CancelRefresh()
		{
			Invoker.Method(this, "CancelRefresh", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Refresh()
		{
			object returnValue = Invoker.MethodReturn(this, "Refresh", null);
			return (bool)returnValue;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Refresh(object backgroundQuery)
		{
			object[] paramArray = new object[1];
			paramArray[0] = backgroundQuery;
			object returnValue = Invoker.MethodReturn(this, "Refresh", paramArray);
			return (bool)returnValue;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void ResetTimer()
		{
			Invoker.Method(this, "ResetTimer", null);
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