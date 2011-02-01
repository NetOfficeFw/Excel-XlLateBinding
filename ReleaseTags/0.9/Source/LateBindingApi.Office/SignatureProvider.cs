//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class SignatureProvider : COMObject
	{
		#region Construction

		public SignatureProvider(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SignatureProvider(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SignatureProvider(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SignatureProvider()
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.stdole.Picture GenerateSignatureLineImage(LateBindingApi.Office.Enums.SignatureLineImage siglnimg, LateBindingApi.Office.SignatureSetup psigsetup, LateBindingApi.Office.SignatureInfo psiginfo, object xmlDsigStream)
		{
			object[] paramArray = new object[4];
			paramArray[0] = siglnimg;
			paramArray.SetValue(psigsetup,1);
			paramArray.SetValue(psiginfo,2);
			paramArray[3] = xmlDsigStream;
			object returnValue = Invoker.MethodReturn(this, "GenerateSignatureLineImage", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.stdole.Picture newClass = new LateBindingApi.stdole.Picture(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public void ShowSignatureSetup(object parentWindow, LateBindingApi.Office.SignatureSetup psigsetup)
		{
			object[] paramArray = new object[2];
			paramArray[0] = parentWindow;
			paramArray.SetValue(psigsetup,1);
			Invoker.Method(this, "ShowSignatureSetup", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void ShowSigningCeremony(object parentWindow, LateBindingApi.Office.SignatureSetup psigsetup, LateBindingApi.Office.SignatureInfo psiginfo)
		{
			object[] paramArray = new object[3];
			paramArray[0] = parentWindow;
			paramArray.SetValue(psigsetup,1);
			paramArray.SetValue(psiginfo,2);
			Invoker.Method(this, "ShowSigningCeremony", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void SignXmlDsig(object queryContinue, LateBindingApi.Office.SignatureSetup psigsetup, LateBindingApi.Office.SignatureInfo psiginfo, object xmlDsigStream)
		{
			object[] paramArray = new object[4];
			paramArray[0] = queryContinue;
			paramArray.SetValue(psigsetup,1);
			paramArray.SetValue(psiginfo,2);
			paramArray[3] = xmlDsigStream;
			Invoker.Method(this, "SignXmlDsig", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void NotifySignatureAdded(object parentWindow, LateBindingApi.Office.SignatureSetup psigsetup, LateBindingApi.Office.SignatureInfo psiginfo)
		{
			object[] paramArray = new object[3];
			paramArray[0] = parentWindow;
			paramArray.SetValue(psigsetup,1);
			paramArray.SetValue(psiginfo,2);
			Invoker.Method(this, "NotifySignatureAdded", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void VerifyXmlDsig(object queryContinue, LateBindingApi.Office.SignatureSetup psigsetup, LateBindingApi.Office.SignatureInfo psiginfo, object xmlDsigStream, LateBindingApi.Office.Enums.ContentVerificationResults pcontverres, LateBindingApi.Office.Enums.CertificateVerificationResults pcertverres)
		{
			object[] paramArray = new object[6];
			paramArray[0] = queryContinue;
			paramArray.SetValue(psigsetup,1);
			paramArray.SetValue(psiginfo,2);
			paramArray[3] = xmlDsigStream;
			paramArray.SetValue(pcontverres,4);
			paramArray.SetValue(pcertverres,5);
			Invoker.Method(this, "VerifyXmlDsig", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void ShowSignatureDetails(object parentWindow, LateBindingApi.Office.SignatureSetup psigsetup, LateBindingApi.Office.SignatureInfo psiginfo, object xmlDsigStream, LateBindingApi.Office.Enums.ContentVerificationResults pcontverres, LateBindingApi.Office.Enums.CertificateVerificationResults pcertverres)
		{
			object[] paramArray = new object[6];
			paramArray[0] = parentWindow;
			paramArray.SetValue(psigsetup,1);
			paramArray.SetValue(psiginfo,2);
			paramArray[3] = xmlDsigStream;
			paramArray.SetValue(pcontverres,4);
			paramArray.SetValue(pcertverres,5);
			Invoker.Method(this, "ShowSignatureDetails", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public COMVariant GetProviderDetail(LateBindingApi.Office.Enums.SignatureProviderDetail sigprovdet)
		{
			object[] paramArray = new object[1];
			paramArray[0] = sigprovdet;
			object returnValue = Invoker.MethodReturn(this, "GetProviderDetail", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("OF12","OF14")]
		public COMVariant HashStream(object queryContinue, object stream)
		{
			object[] paramArray = new object[2];
			paramArray[0] = queryContinue;
			paramArray[1] = stream;
			object returnValue = Invoker.MethodReturn(this, "HashStream", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
