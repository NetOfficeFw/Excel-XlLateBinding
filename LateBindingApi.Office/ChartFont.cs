//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class ChartFont : COMObject
	{
		#region Construction

		public ChartFont(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ChartFont(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ChartFont(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ChartFont()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public COMVariant Background
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Background");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Background", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Bold
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Bold");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Bold", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Color
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Color");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Color", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant ColorIndex
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ColorIndex");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "ColorIndex", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant FontStyle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FontStyle");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "FontStyle", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Italic
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Italic");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Italic", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Name", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant OutlineFont
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OutlineFont");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "OutlineFont", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Shadow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Shadow");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Shadow", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Size
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Size");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Size", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant StrikeThrough
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "StrikeThrough");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "StrikeThrough", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Subscript
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Subscript");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Subscript", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Superscript
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Superscript");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Superscript", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public COMVariant Underline
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Underline");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Underline", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public COMObject Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF14")]
		public Int32 Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}