//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class TextEffectFormat : LateBindingApi.Office._IMsoDispObj
	{
		#region Construction

		public TextEffectFormat(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public TextEffectFormat(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public TextEffectFormat(COMObject replacedObject) : base(replacedObject)
		{
		}

		public TextEffectFormat()
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
		public LateBindingApi.Office.Enums.MsoTextEffectAlignment Alignment
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Alignment");
				return (LateBindingApi.Office.Enums.MsoTextEffectAlignment)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Alignment", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoTriState FontBold
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FontBold");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FontBold", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoTriState FontItalic
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FontItalic");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FontItalic", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string FontName
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FontName");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FontName", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double FontSize
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FontSize");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FontSize", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoTriState KernedPairs
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "KernedPairs");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "KernedPairs", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoTriState NormalizedHeight
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "NormalizedHeight");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "NormalizedHeight", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoPresetTextEffectShape PresetShape
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PresetShape");
				return (LateBindingApi.Office.Enums.MsoPresetTextEffectShape)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PresetShape", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoPresetTextEffect PresetTextEffect
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PresetTextEffect");
				return (LateBindingApi.Office.Enums.MsoPresetTextEffect)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PresetTextEffect", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Office.Enums.MsoTriState RotatedChars
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RotatedChars");
				return (LateBindingApi.Office.Enums.MsoTriState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RotatedChars", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string Text
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Text");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Text", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Tracking
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Tracking");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Tracking", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void ToggleVerticalText()
		{
			Invoker.Method(this, "ToggleVerticalText", null);
		}

		#endregion

	}
}
