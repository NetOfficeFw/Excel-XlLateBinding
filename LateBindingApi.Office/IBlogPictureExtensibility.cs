//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class IBlogPictureExtensibility : COMObject
	{
		#region Construction

		public IBlogPictureExtensibility(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IBlogPictureExtensibility(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IBlogPictureExtensibility(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IBlogPictureExtensibility()
		{
		}

		#endregion

		#region Properties

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public void BlogPictureProviderProperties(string blogPictureProvider, string friendlyName)
		{
			object[] paramArray = new object[2];
			paramArray.SetValue(blogPictureProvider,0);
			paramArray.SetValue(friendlyName,1);
			Invoker.Method(this, "BlogPictureProviderProperties", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void CreatePictureAccount(string account, string blogProvider, Int32 parentWindow, COMObject document)
		{
			object[] paramArray = new object[4];
			paramArray[0] = account;
			paramArray[1] = blogProvider;
			paramArray[2] = parentWindow;
			paramArray[3] = document;
			Invoker.Method(this, "CreatePictureAccount", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void PublishPicture(string account, Int32 parentWindow, COMObject document, object image, string pictureURI, Int32 imageType)
		{
			object[] paramArray = new object[6];
			paramArray[0] = account;
			paramArray[1] = parentWindow;
			paramArray[2] = document;
			paramArray[3] = image;
			paramArray.SetValue(pictureURI,4);
			paramArray[5] = imageType;
			Invoker.Method(this, "PublishPicture", paramArray);
		}

		#endregion

	}
}