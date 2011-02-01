//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14")]
	public class Graphic : COMObject
	{
		#region Construction

		public Graphic(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Graphic(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Graphic(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Graphic()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
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

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double Brightness
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Brightness");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Brightness", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Office.Enums.MsoPictureColorType ColorType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ColorType");
				return (LateBindingApi.Office.Enums.MsoPictureColorType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ColorType", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double Contrast
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Contrast");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Contrast", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double CropBottom
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CropBottom");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CropBottom", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double CropLeft
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CropLeft");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CropLeft", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double CropRight
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CropRight");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CropRight", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double CropTop
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CropTop");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CropTop", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public string Filename
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Filename");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Filename", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double Height
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Height");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Height", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Office.Enums.MsoTriState LockAspectRatio
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LockAspectRatio");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LockAspectRatio", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public Double Width
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Width");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Width", value);
			}						
		}


		#endregion

		#region Methods

		#endregion

	}
}
