//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class Window : COMObject
	{
		#region Construction

		public Window(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public Window(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public Window(COMObject replacedObject) : base(replacedObject)
		{
		}

		public Window()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.VBE VBE
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VBE");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.VBE newClass = new LateBindingApi.VBIDE.VBE(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Windows Collection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Collection");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.Windows newClass = new LateBindingApi.VBIDE.Windows(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public string Caption
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Caption");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("VBE")]
		public bool Visible
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Visible");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Visible", value);
			}						
		}


		[SupportByLibrary("VBE")]
		public Int32 Left
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Left");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Left", value);
			}						
		}


		[SupportByLibrary("VBE")]
		public Int32 Top
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Top");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Top", value);
			}						
		}


		[SupportByLibrary("VBE")]
		public Int32 Width
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Width");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Width", value);
			}						
		}


		[SupportByLibrary("VBE")]
		public Int32 Height
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Height");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Height", value);
			}						
		}


		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Enums.Vbext_WindowState WindowState
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WindowState");
				return (LateBindingApi.VBIDE.Enums.Vbext_WindowState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WindowState", value);
			}						
		}


		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Enums.Vbext_WindowType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.VBIDE.Enums.Vbext_WindowType)returnValue;
			}
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.LinkedWindows LinkedWindows
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LinkedWindows");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.LinkedWindows newClass = new LateBindingApi.VBIDE.LinkedWindows(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Window LinkedWindowFrame
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LinkedWindowFrame");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.Window newClass = new LateBindingApi.VBIDE.Window(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public Int32 HWnd
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HWnd");
				return (Int32)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("VBE")]
		public void Close()
		{
			Invoker.Method(this, "Close", null);
		}

		[SupportByLibrary("VBE")]
		public void SetFocus()
		{
			Invoker.Method(this, "SetFocus", null);
		}

		[SupportByLibrary("VBE")]
		public void SetKind(LateBindingApi.VBIDE.Enums.Vbext_WindowType eKind)
		{
			object[] paramArray = new object[1];
			paramArray[0] = eKind;
			Invoker.Method(this, "SetKind", paramArray);
		}

		[SupportByLibrary("VBE")]
		public void Detach()
		{
			Invoker.Method(this, "Detach", null);
		}

		[SupportByLibrary("VBE")]
		public void Attach(Int32 lWindowHandle)
		{
			object[] paramArray = new object[1];
			paramArray[0] = lWindowHandle;
			Invoker.Method(this, "Attach", paramArray);
		}

		#endregion

	}
}