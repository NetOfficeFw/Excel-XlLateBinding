using System;

namespace LateBindingApi.Excel.Enums
{
	public enum MsoBarProtection
	{
		msoBarNoProtection = 0,
		msoBarNoCustomize = 1,
		msoBarNoResize = 2,
		msoBarNoMove = 4,
		msoBarNoChangeVisible = 8,
		msoBarNoChangeDock = 16,
		msoBarNoVerticalDock = 32,
		msoBarNoHorizontalDock = 64
	}
}