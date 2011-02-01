//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class IMsoTickLabels : COMObject
	{
		#region Construction

		public IMsoTickLabels(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IMsoTickLabels(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IMsoTickLabels(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IMsoTickLabels()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.ChartFont Font
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Font");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.ChartFont newClass = new LateBindingApi.Office.ChartFont(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.XlTickLabelOrientation Orientation
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Orientation");
				return (LateBindingApi.Office.Enums.XlTickLabelOrientation)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Orientation", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
		public Int32 Depth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Depth");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.IMsoChartFormat Format
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Format");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IMsoChartFormat newClass = new LateBindingApi.Office.IMsoChartFormat(this, returnValue);
				return newClass;
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

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public COMVariant Delete()
		{
			object returnValue = Invoker.MethodReturn(this, "Delete", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("OF12","OF14")]
		public COMVariant Select()
		{
			object returnValue = Invoker.MethodReturn(this, "Select", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
