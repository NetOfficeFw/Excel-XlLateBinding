//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF14")]
	public class PictureEffect : _IMsoDispObj
	{
		#region Construction

		public PictureEffect(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PictureEffect(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PictureEffect(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PictureEffect()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF14")]
		public LateBindingApi.Office.Enums.MsoPictureEffectType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Office.Enums.MsoPictureEffectType)returnValue;
			}
		}

		[SupportByLibrary("OF14")]
		public Int32 Position
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Position");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Position", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public LateBindingApi.Office.EffectParameters EffectParameters
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "EffectParameters");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.EffectParameters newClass = new LateBindingApi.Office.EffectParameters(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF14")]
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


		#endregion

		#region Methods

		[SupportByLibrary("OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		#endregion

	}
}
