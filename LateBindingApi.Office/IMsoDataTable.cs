//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class IMsoDataTable : COMObject
	{
		#region Construction

		public IMsoDataTable(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IMsoDataTable(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IMsoDataTable(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IMsoDataTable()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public bool ShowLegendKey
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ShowLegendKey");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ShowLegendKey", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public bool HasBorderHorizontal
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasBorderHorizontal");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasBorderHorizontal", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public bool HasBorderVertical
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasBorderVertical");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasBorderVertical", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public bool HasBorderOutline
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasBorderOutline");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HasBorderOutline", value);
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
		public void Select()
		{
			Invoker.Method(this, "Select", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		#endregion

	}
}
