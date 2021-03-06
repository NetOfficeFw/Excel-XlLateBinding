//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class ChartArea : COMObject
	{
		#region Construction

		public ChartArea(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ChartArea(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ChartArea(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ChartArea()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Border Border
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Border");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Border newClass = new LateBindingApi.Excel.Border(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Font Font
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Font");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Font newClass = new LateBindingApi.Excel.Font(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Shadow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Shadow");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Shadow", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Interior Interior
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Interior");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Interior newClass = new LateBindingApi.Excel.Interior(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.ChartFillFormat Fill
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Fill");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.ChartFillFormat newClass = new LateBindingApi.Excel.ChartFillFormat(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Left
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Left");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Left", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Top
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Top");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Top", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant AutoScaleFont
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AutoScaleFont");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "AutoScaleFont", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.ChartFormat Format
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Format");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.ChartFormat newClass = new LateBindingApi.Excel.ChartFormat(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL14")]
		public bool RoundedCorners
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RoundedCorners");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RoundedCorners", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Select()
		{
			object returnValue = Invoker.MethodReturn(this, "Select", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Clear()
		{
			object returnValue = Invoker.MethodReturn(this, "Clear", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ClearContents()
		{
			object returnValue = Invoker.MethodReturn(this, "ClearContents", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Copy()
		{
			object returnValue = Invoker.MethodReturn(this, "Copy", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ClearFormats()
		{
			object returnValue = Invoker.MethodReturn(this, "ClearFormats", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
