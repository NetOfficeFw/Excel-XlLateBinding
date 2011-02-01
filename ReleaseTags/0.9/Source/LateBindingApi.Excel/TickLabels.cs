//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class TickLabels : COMObject
	{
		#region Construction

		public TickLabels(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public TickLabels(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public TickLabels(COMObject replacedObject) : base(replacedObject)
		{
		}

		public TickLabels()
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
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string NumberFormat
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NumberFormat");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "NumberFormat", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool NumberFormatLinked
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NumberFormatLinked");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "NumberFormatLinked", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant NumberFormatLocal
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NumberFormatLocal");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "NumberFormatLocal", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlTickLabelOrientation Orientation
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Orientation");
				return (LateBindingApi.Excel.Enums.XlTickLabelOrientation)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Orientation", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ReadingOrder
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ReadingOrder");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ReadingOrder", value);
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Depth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Depth");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Offset
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Offset");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Offset", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Alignment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Alignment");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Alignment", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public bool MultiLevel
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MultiLevel");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "MultiLevel", value);
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

		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Delete()
		{
			object returnValue = Invoker.MethodReturn(this, "Delete", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Select()
		{
			object returnValue = Invoker.MethodReturn(this, "Select", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
