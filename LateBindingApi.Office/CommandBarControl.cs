//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class CommandBarControl : _IMsoOleAccDispObj
	{
		#region Construction

		public CommandBarControl(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public CommandBarControl(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public CommandBarControl(COMObject replacedObject) : base(replacedObject)
		{
		}

		public CommandBarControl()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool BeginGroup
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BeginGroup");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "BeginGroup", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool BuiltIn
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BuiltIn");
				return (bool)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Caption
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Caption");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Caption", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMObject Control
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Control");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string DescriptionText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DescriptionText");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DescriptionText", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool Enabled
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Enabled");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Enabled", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
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


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 HelpContextId
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HelpContextId");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HelpContextId", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string HelpFile
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HelpFile");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HelpFile", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 Id
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Id");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 Index
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Index");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 InstanceId
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InstanceId");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 Left
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Left");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoControlOLEUsage OLEUsage
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OLEUsage");
				return (LateBindingApi.Office.Enums.MsoControlOLEUsage)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OLEUsage", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string OnAction
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OnAction");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OnAction", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBar Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.CommandBar newClass = new LateBindingApi.Office.CommandBar(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Parameter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parameter");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Parameter", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 Priority
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Priority");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Priority", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Tag
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Tag");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Tag", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string TooltipText
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TooltipText");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TooltipText", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 Top
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Top");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoControlType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Office.Enums.MsoControlType)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
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


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
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


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool IsPriorityDropped
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "IsPriorityDropped");
				return (bool)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBarControl Copy()
		{
			object returnValue = Invoker.MethodReturn(this, "Copy", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CommandBarControl newClass = new LateBindingApi.Office.CommandBarControl(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBarControl Copy(object bar, object before)
		{
			object[] paramArray = new object[2];
			paramArray[0] = bar;
			paramArray[1] = before;
			object returnValue = Invoker.MethodReturn(this, "Copy", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CommandBarControl newClass = new LateBindingApi.Office.CommandBarControl(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Delete(object temporary)
		{
			object[] paramArray = new object[1];
			paramArray[0] = temporary;
			Invoker.Method(this, "Delete", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Execute()
		{
			Invoker.Method(this, "Execute", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBarControl Move()
		{
			object returnValue = Invoker.MethodReturn(this, "Move", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CommandBarControl newClass = new LateBindingApi.Office.CommandBarControl(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.CommandBarControl Move(object bar, object before)
		{
			object[] paramArray = new object[2];
			paramArray[0] = bar;
			paramArray[1] = before;
			object returnValue = Invoker.MethodReturn(this, "Move", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.CommandBarControl newClass = new LateBindingApi.Office.CommandBarControl(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reset()
		{
			Invoker.Method(this, "Reset", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void SetFocus()
		{
			Invoker.Method(this, "SetFocus", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reserved1()
		{
			Invoker.Method(this, "Reserved1", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reserved2()
		{
			Invoker.Method(this, "Reserved2", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reserved3()
		{
			Invoker.Method(this, "Reserved3", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reserved4()
		{
			Invoker.Method(this, "Reserved4", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reserved5()
		{
			Invoker.Method(this, "Reserved5", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reserved6()
		{
			Invoker.Method(this, "Reserved6", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Reserved7()
		{
			Invoker.Method(this, "Reserved7", null);
		}

		#endregion

	}
}
