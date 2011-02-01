//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class IMsoPlotArea : COMObject
	{
		#region Construction

		public IMsoPlotArea(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IMsoPlotArea(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IMsoPlotArea(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IMsoPlotArea()
		{
		}

		#endregion

		#region Properties

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
		public LateBindingApi.Office.IMsoBorder Border
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Border");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IMsoBorder newClass = new LateBindingApi.Office.IMsoBorder(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.IMsoInterior Interior
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Interior");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IMsoInterior newClass = new LateBindingApi.Office.IMsoInterior(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.ChartFillFormat Fill
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Fill");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.ChartFillFormat newClass = new LateBindingApi.Office.ChartFillFormat(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.XlChartElementPosition Position
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Position");
				return (LateBindingApi.Office.Enums.XlChartElementPosition)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Position", value);
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
		public COMVariant Select()
		{
			object returnValue = Invoker.MethodReturn(this, "Select", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("OF12","OF14")]
		public COMVariant ClearFormats()
		{
			object returnValue = Invoker.MethodReturn(this, "ClearFormats", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
