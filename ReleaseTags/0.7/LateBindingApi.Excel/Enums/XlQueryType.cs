using System;

namespace LateBindingApi.Excel.Enums
{
	public enum XlQueryType
	{
		xlODBCQuery = 1,
		xlDAORecordset = 2,
		xlWebQuery = 4,
		xlOLEDBQuery = 5,
		xlTextImport = 6,
		xlADORecordset = 7
	}
}