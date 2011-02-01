//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14")]
	public class ILicWizExternal : COMObject
	{
		#region Construction

		public ILicWizExternal(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ILicWizExternal(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ILicWizExternal(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ILicWizExternal()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 Context
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Context");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject Validator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Validator");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public COMObject LicAgent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LicAgent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string CountryInfo
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CountryInfo");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 WizardVisible
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WizardVisible");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WizardVisible", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string WizardTitle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WizardTitle");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WizardTitle", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 AnimationEnabled
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AnimationEnabled");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 CurrentHelpId
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CurrentHelpId");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CurrentHelpId", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string OfficeOnTheWebUrl
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OfficeOnTheWebUrl");
				return (string)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void PrintHtmlDocument(object punkHtmlDoc)
		{
			object[] paramArray = new object[1];
			paramArray[0] = punkHtmlDoc;
			Invoker.Method(this, "PrintHtmlDocument", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void InvokeDateTimeApplet()
		{
			Invoker.Method(this, "InvokeDateTimeApplet", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string FormatDate(object date, string pFormat)
		{
			object[] paramArray = new object[2];
			paramArray[0] = date;
			paramArray[1] = pFormat;
			object returnValue = Invoker.MethodReturn(this, "FormatDate", paramArray);
			return (string)returnValue;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void ShowHelp()
		{
			Invoker.Method(this, "ShowHelp", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void ShowHelp(object pvarId)
		{
			object[] paramArray = new object[1];
			paramArray.SetValue(pvarId,0);
			Invoker.Method(this, "ShowHelp", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void Terminate()
		{
			Invoker.Method(this, "Terminate", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void DisableVORWReminder(Int32 bPC)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bPC;
			Invoker.Method(this, "DisableVORWReminder", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public string SaveReceipt(string bstrReceipt)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrReceipt;
			object returnValue = Invoker.MethodReturn(this, "SaveReceipt", paramArray);
			return (string)returnValue;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void OpenInDefaultBrowser(string bstrUrl)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrUrl;
			Invoker.Method(this, "OpenInDefaultBrowser", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 MsoAlert(string bstrText, string bstrButtons, string bstrIcon)
		{
			object[] paramArray = new object[3];
			paramArray[0] = bstrText;
			paramArray[1] = bstrButtons;
			paramArray[2] = bstrIcon;
			object returnValue = Invoker.MethodReturn(this, "MsoAlert", paramArray);
			return (Int32)returnValue;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 DepositPidKey(string bstrKey, Int32 fMORW)
		{
			object[] paramArray = new object[2];
			paramArray[0] = bstrKey;
			paramArray[1] = fMORW;
			object returnValue = Invoker.MethodReturn(this, "DepositPidKey", paramArray);
			return (Int32)returnValue;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void WriteLog(string bstrMessage)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrMessage;
			Invoker.Method(this, "WriteLog", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void ResignDpc(string bstrProductCode)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrProductCode;
			Invoker.Method(this, "ResignDpc", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void ResetPID()
		{
			Invoker.Method(this, "ResetPID", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void SetDialogSize(Int32 dx, Int32 dy)
		{
			object[] paramArray = new object[2];
			paramArray[0] = dx;
			paramArray[1] = dy;
			Invoker.Method(this, "SetDialogSize", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 VerifyClock(Int32 lMode)
		{
			object[] paramArray = new object[1];
			paramArray[0] = lMode;
			object returnValue = Invoker.MethodReturn(this, "VerifyClock", paramArray);
			return (Int32)returnValue;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void SortSelectOptions(COMObject pdispSelect)
		{
			object[] paramArray = new object[1];
			paramArray[0] = pdispSelect;
			Invoker.Method(this, "SortSelectOptions", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public void InternetDisconnect()
		{
			Invoker.Method(this, "InternetDisconnect", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14")]
		public Int32 GetConnectedState()
		{
			object returnValue = Invoker.MethodReturn(this, "GetConnectedState", null);
			return (Int32)returnValue;
		}

		#endregion

	}
}
