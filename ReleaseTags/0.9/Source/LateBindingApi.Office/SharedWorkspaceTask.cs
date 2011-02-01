//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF11","OF12","OF14")]
	public class SharedWorkspaceTask : _IMsoDispObj
	{
		#region Construction

		public SharedWorkspaceTask(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SharedWorkspaceTask(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SharedWorkspaceTask(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SharedWorkspaceTask()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF11","OF12","OF14")]
		public string Title
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Title");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Title", value);
			}						
		}


		[SupportByLibrary("OF11","OF12","OF14")]
		public string AssignedTo
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AssignedTo");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AssignedTo", value);
			}						
		}


		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoSharedWorkspaceTaskStatus Status
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Status");
				return (LateBindingApi.Office.Enums.MsoSharedWorkspaceTaskStatus)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Status", value);
			}						
		}


		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoSharedWorkspaceTaskPriority Priority
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Priority");
				return (LateBindingApi.Office.Enums.MsoSharedWorkspaceTaskPriority)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Priority", value);
			}						
		}


		[SupportByLibrary("OF11","OF12","OF14")]
		public string Description
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Description");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Description", value);
			}						
		}


		[SupportByLibrary("OF11","OF12","OF14")]
		public COMVariant DueDate
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DueDate");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "DueDate", value);
			}						
		}


		[SupportByLibrary("OF11","OF12","OF14")]
		public string CreatedBy
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CreatedBy");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMVariant CreatedDate
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CreatedDate");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public string ModifiedBy
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ModifiedBy");
				return (string)returnValue;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMVariant ModifiedDate
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ModifiedDate");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF11","OF12","OF14")]
		public void Save()
		{
			Invoker.Method(this, "Save", null);
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		#endregion

	}
}