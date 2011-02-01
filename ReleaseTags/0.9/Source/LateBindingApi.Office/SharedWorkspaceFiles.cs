//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF11","OF12","OF14")]
	public class SharedWorkspaceFiles : _IMsoDispObj
	{
		#region Construction

		public SharedWorkspaceFiles(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public SharedWorkspaceFiles(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public SharedWorkspaceFiles(COMObject replacedObject) : base(replacedObject)
		{
		}

		public SharedWorkspaceFiles()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF11","OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.SharedWorkspaceFile returnClass = new LateBindingApi.Office.SharedWorkspaceFile (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.SharedWorkspaceFile this[Int32 index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.PropertyGet(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.SharedWorkspaceFile newClass = new LateBindingApi.Office.SharedWorkspaceFile(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
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

		[SupportByLibrary("OF11","OF12","OF14")]
		public bool ItemCountExceeded
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ItemCountExceeded");
				return (bool)returnValue;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.SharedWorkspaceFile Add(string fileName)
		{
			object[] paramArray = new object[1];
			paramArray[0] = fileName;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.SharedWorkspaceFile newClass = new LateBindingApi.Office.SharedWorkspaceFile(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF11","OF12","OF14")]
		public LateBindingApi.Office.SharedWorkspaceFile Add(string fileName, object parentFolder, object overwriteIfFileAlreadyExists, object keepInSync)
		{
			object[] paramArray = new object[4];
			paramArray[0] = fileName;
			paramArray[1] = parentFolder;
			paramArray[2] = overwriteIfFileAlreadyExists;
			paramArray[3] = keepInSync;
			object returnValue = Invoker.MethodReturn(this, "Add", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.SharedWorkspaceFile newClass = new LateBindingApi.Office.SharedWorkspaceFile(this, returnValue);
			return newClass;
		}

		#endregion

	}
}