//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class SignatureInfo : _IMsoDispObj
	{
		#region Construction

		public SignatureInfo(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SignatureInfo(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SignatureInfo(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SignatureInfo()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public bool ReadOnly
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ReadOnly");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string SignatureProvider
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SignatureProvider");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string SignatureText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SignatureText");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SignatureText", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.stdole.Picture SignatureImage
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SignatureImage");
				if(null == returnValue)
					return null;
				LateBindingApi.stdole.Picture newClass = new LateBindingApi.stdole.Picture(this, returnValue);
				return newClass;
			}
			set
			{
				Invoker.PropertySet(this, "SignatureImage", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public string SignatureComment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SignatureComment");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SignatureComment", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.ContentVerificationResults ContentVerificationResults
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ContentVerificationResults");
				return (LateBindingApi.Office.Enums.ContentVerificationResults)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.CertificateVerificationResults CertificateVerificationResults
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CertificateVerificationResults");
				return (LateBindingApi.Office.Enums.CertificateVerificationResults)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public bool IsValid
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsValid");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public bool IsCertificateExpired
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsCertificateExpired");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public bool IsCertificateRevoked
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsCertificateRevoked");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public bool IsCertificateUntrusted
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsCertificateUntrusted");
				return (bool)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public COMVariant GetSignatureDetail(LateBindingApi.Office.Enums.SignatureDetail sigdet)
		{
			object[] paramArray = new object[1];
			paramArray[0] = sigdet;
			object returnValue = Invoker.MethodReturn(this, "GetSignatureDetail", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("OF12","OF14")]
		public COMVariant GetCertificateDetail(LateBindingApi.Office.Enums.CertificateDetail certdet)
		{
			object[] paramArray = new object[1];
			paramArray[0] = certdet;
			object returnValue = Invoker.MethodReturn(this, "GetCertificateDetail", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("OF12","OF14")]
		public void ShowSignatureCertificate(object parentWindow)
		{
			object[] paramArray = new object[1];
			paramArray[0] = parentWindow;
			Invoker.Method(this, "ShowSignatureCertificate", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void SelectSignatureCertificate(object parentWindow)
		{
			object[] paramArray = new object[1];
			paramArray[0] = parentWindow;
			Invoker.Method(this, "SelectSignatureCertificate", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void SelectCertificateDetailByThumbprint(string bstrThumbprint)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrThumbprint;
			Invoker.Method(this, "SelectCertificateDetailByThumbprint", paramArray);
		}

		#endregion

	}
}
