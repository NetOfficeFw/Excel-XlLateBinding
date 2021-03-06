//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.VBIDE
{
	[SupportByLibrary("VBE")]
	public class _VBComponent_Old : COMObject
	{
		#region Construction

		public _VBComponent_Old(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _VBComponent_Old(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _VBComponent_Old(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _VBComponent_Old()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("VBE")]
		public bool Saved
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Saved");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("VBE")]
		public string Name
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Name");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Name", value);
			}						
		}


		[SupportByLibrary("VBE")]
		public COMObject Designer
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Designer");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.CodeModule CodeModule
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CodeModule");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.CodeModule newClass = new LateBindingApi.VBIDE.CodeModule(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Enums.Vbext_ComponentType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.VBIDE.Enums.Vbext_ComponentType)returnValue;
			}
		}

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
		public LateBindingApi.VBIDE.VBComponents Collection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Collection");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.VBComponents newClass = new LateBindingApi.VBIDE.VBComponents(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("VBE")]
		public bool HasOpenDesigner
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HasOpenDesigner");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Properties Properties
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Properties");
				if(null == returnValue)
					return null;
				LateBindingApi.VBIDE.Properties newClass = new LateBindingApi.VBIDE.Properties(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("VBE")]
		public void Export(string fileName)
		{
			object[] paramArray = new object[1];
			paramArray[0] = fileName;
			Invoker.Method(this, "Export", paramArray);
		}

		[SupportByLibrary("VBE")]
		public LateBindingApi.VBIDE.Window DesignerWindow()
		{
			object returnValue = Invoker.MethodReturn(this, "DesignerWindow", null);
			if(null == returnValue)
				return null;
			LateBindingApi.VBIDE.Window newClass = new LateBindingApi.VBIDE.Window(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("VBE")]
		public void Activate()
		{
			Invoker.Method(this, "Activate", null);
		}

		#endregion

	}
}
