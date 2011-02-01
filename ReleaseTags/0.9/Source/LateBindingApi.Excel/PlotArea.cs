//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class PlotArea : COMObject
	{
		#region Construction

		public PlotArea(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PlotArea(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PlotArea(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PlotArea()
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
		public Double InsideLeft
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InsideLeft");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InsideLeft", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double InsideTop
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InsideTop");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InsideTop", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double InsideWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InsideWidth");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InsideWidth", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double InsideHeight
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InsideHeight");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "InsideHeight", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public Double _InsideLeft
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_InsideLeft");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public Double _InsideTop
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_InsideTop");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public Double _InsideWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_InsideWidth");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public Double _InsideHeight
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_InsideHeight");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlChartElementPosition Position
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Position");
				return (LateBindingApi.Excel.Enums.XlChartElementPosition)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Position", value);
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
		public COMVariant Select()
		{
			object returnValue = Invoker.MethodReturn(this, "Select", null);
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
