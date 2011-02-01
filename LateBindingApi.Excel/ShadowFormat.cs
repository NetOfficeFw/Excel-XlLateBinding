//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class ShadowFormat : LateBindingApi.Office._IMsoDispObj
	{
		#region Construction

		public ShadowFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public ShadowFormat(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public ShadowFormat(COMObject replacedObject) : base(replacedObject)
		{
		}

		public ShadowFormat()
		{
		}

		#endregion

		#region Properties

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
		public LateBindingApi.Excel.ColorFormat ForeColor
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ForeColor");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.ColorFormat newClass = new LateBindingApi.Excel.ColorFormat(this, returnValue);
				return newClass;
			}
			set
			{
				Invoker.PropertySet(this, "ForeColor", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoTriState Obscured
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Obscured");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Obscured", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double OffsetX
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OffsetX");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OffsetX", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double OffsetY
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OffsetY");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OffsetY", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Transparency
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Transparency");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Transparency", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoShadowType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Office.Enums.MsoShadowType)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Type", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoTriState Visible
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Visible");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Visible", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Office.Enums.MsoShadowStyle Style
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Style");
				return (LateBindingApi.Office.Enums.MsoShadowStyle)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Style", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public Double Blur
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Blur");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Blur", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public Double Size
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Size");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Size", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Office.Enums.MsoTriState RotateWithShape
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RotateWithShape");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RotateWithShape", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void IncrementOffsetX(Double increment)
		{
			object[] paramArray = new object[1];
			paramArray[0] = increment;
			Invoker.Method(this, "IncrementOffsetX", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void IncrementOffsetY(Double increment)
		{
			object[] paramArray = new object[1];
			paramArray[0] = increment;
			Invoker.Method(this, "IncrementOffsetY", paramArray);
		}

		#endregion

	}
}