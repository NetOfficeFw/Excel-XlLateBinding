//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class ChartGroup : COMObject
	{
		#region Construction

		public ChartGroup(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ChartGroup(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ChartGroup(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ChartGroup()
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
		public LateBindingApi.Excel.Enums.XlAxisGroup AxisGroup
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AxisGroup");
				return (LateBindingApi.Excel.Enums.XlAxisGroup)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AxisGroup", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 DoughnutHoleSize
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DoughnutHoleSize");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DoughnutHoleSize", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.DownBars DownBars
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DownBars");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.DownBars newClass = new LateBindingApi.Excel.DownBars(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.DropLines DropLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DropLines");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.DropLines newClass = new LateBindingApi.Excel.DropLines(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 FirstSliceAngle
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FirstSliceAngle");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FirstSliceAngle", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 GapWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "GapWidth");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "GapWidth", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool HasDropLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasDropLines");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasDropLines", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool HasHiLoLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasHiLoLines");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasHiLoLines", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool HasRadarAxisLabels
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasRadarAxisLabels");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasRadarAxisLabels", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool HasSeriesLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasSeriesLines");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasSeriesLines", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool HasUpDownBars
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasUpDownBars");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasUpDownBars", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.HiLoLines HiLoLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HiLoLines");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.HiLoLines newClass = new LateBindingApi.Excel.HiLoLines(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Index
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Index");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Overlap
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Overlap");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Overlap", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.TickLabels RadarAxisLabels
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RadarAxisLabels");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.TickLabels newClass = new LateBindingApi.Excel.TickLabels(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.SeriesLines SeriesLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SeriesLines");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.SeriesLines newClass = new LateBindingApi.Excel.SeriesLines(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 SubType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SubType");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SubType", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Type", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.UpBars UpBars
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "UpBars");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.UpBars newClass = new LateBindingApi.Excel.UpBars(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool VaryByCategories
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VaryByCategories");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "VaryByCategories", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlSizeRepresents SizeRepresents
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SizeRepresents");
				return (LateBindingApi.Excel.Enums.XlSizeRepresents)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SizeRepresents", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 BubbleScale
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BubbleScale");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "BubbleScale", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool ShowNegativeBubbles
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShowNegativeBubbles");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ShowNegativeBubbles", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlChartSplitType SplitType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SplitType");
				return (LateBindingApi.Excel.Enums.XlChartSplitType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SplitType", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SplitValue
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SplitValue");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "SplitValue", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 SecondPlotSize
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SecondPlotSize");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SecondPlotSize", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Has3DShading
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Has3DShading");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Has3DShading", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject SeriesCollection()
		{
			object returnValue = Invoker.MethodReturn(this, "SeriesCollection", null);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject SeriesCollection(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			object returnValue = Invoker.MethodReturn(this, "SeriesCollection", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
