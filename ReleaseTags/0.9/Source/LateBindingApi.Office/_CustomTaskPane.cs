//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class _CustomTaskPane : COMObject
	{
		#region Construction

		public _CustomTaskPane(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _CustomTaskPane(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _CustomTaskPane(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _CustomTaskPane()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public string Title
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Title");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public COMObject Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public COMObject Window
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Window");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
		public COMObject ContentControl
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ContentControl");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
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


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoCTPDockPosition DockPosition
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DockPosition");
				return (LateBindingApi.Office.Enums.MsoCTPDockPosition)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DockPosition", value);
			}						
		}


		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoCTPDockPositionRestrict DockPositionRestrict
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DockPositionRestrict");
				return (LateBindingApi.Office.Enums.MsoCTPDockPositionRestrict)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DockPositionRestrict", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		#endregion

	}
}
