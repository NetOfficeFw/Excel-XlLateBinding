//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL12","XL14")]
	public class HeaderFooter : COMObject
	{
		#region Construction

		public HeaderFooter(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public HeaderFooter(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public HeaderFooter(COMObject replacedObject) : base(replacedObject)
		{
		}

		public HeaderFooter()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL12","XL14")]
		public string Text
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Text");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Text", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Graphic Picture
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Picture");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Graphic newClass = new LateBindingApi.Excel.Graphic(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
