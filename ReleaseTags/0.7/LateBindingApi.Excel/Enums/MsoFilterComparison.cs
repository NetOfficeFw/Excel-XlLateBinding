using System;

namespace LateBindingApi.Excel.Enums
{
	public enum MsoFilterComparison
	{
		msoFilterComparisonEqual = 0,
		msoFilterComparisonNotEqual = 1,
		msoFilterComparisonLessThan = 2,
		msoFilterComparisonGreaterThan = 3,
		msoFilterComparisonLessThanEqual = 4,
		msoFilterComparisonGreaterThanEqual = 5,
		msoFilterComparisonIsBlank = 6,
		msoFilterComparisonIsNotBlank = 7,
		msoFilterComparisonContains = 8,
		msoFilterComparisonNotContains = 9
	}
}