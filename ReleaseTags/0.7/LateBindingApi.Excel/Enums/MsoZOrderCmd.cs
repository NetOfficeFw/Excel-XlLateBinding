using System;

namespace LateBindingApi.Excel.Enums
{
	public enum MsoZOrderCmd
	{
		msoBringToFront = 0,
		msoSendToBack = 1,
		msoBringForward = 2,
		msoSendBackward = 3,
		msoBringInFrontOfText = 4,
		msoSendBehindText = 5
	}
}