//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class IFind : COMObject
	{
		#region Construction

		public IFind(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public IFind(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public IFind(COMObject replacedObject) : base(replacedObject)
		{
		}

		public IFind()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string SearchPath
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SearchPath");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SearchPath", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
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


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool SubDir
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SubDir");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SubDir", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
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


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Author
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Author");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Author", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Keywords
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Keywords");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Keywords", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string Subject
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Subject");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Subject", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoFileFindOptions Options
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Options");
				return (LateBindingApi.Office.Enums.MsoFileFindOptions)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Options", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool MatchCase
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "MatchCase");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "MatchCase", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
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


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public bool PatternMatch
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PatternMatch");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PatternMatch", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMVariant DateSavedFrom
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DateSavedFrom");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "DateSavedFrom", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMVariant DateSavedTo
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DateSavedTo");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "DateSavedTo", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string SavedBy
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SavedBy");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SavedBy", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMVariant DateCreatedFrom
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DateCreatedFrom");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "DateCreatedFrom", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public COMVariant DateCreatedTo
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DateCreatedTo");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "DateCreatedTo", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoFileFindView View
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "View");
				return (LateBindingApi.Office.Enums.MsoFileFindView)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "View", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoFileFindSortBy SortBy
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SortBy");
				return (LateBindingApi.Office.Enums.MsoFileFindSortBy)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SortBy", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoFileFindListBy ListBy
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListBy");
				return (LateBindingApi.Office.Enums.MsoFileFindListBy)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ListBy", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 SelectedFile
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SelectedFile");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SelectedFile", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.IFoundFiles Results
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Results");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.IFoundFiles newClass = new LateBindingApi.Office.IFoundFiles(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 FileType
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FileType");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FileType", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 Show()
		{
			object returnValue = Invoker.MethodReturn(this, "Show", null);
			return (Int32)returnValue;
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Execute()
		{
			Invoker.Method(this, "Execute", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Load(string bstrQueryName)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrQueryName;
			Invoker.Method(this, "Load", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Save(string bstrQueryName)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrQueryName;
			Invoker.Method(this, "Save", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Delete(string bstrQueryName)
		{
			object[] paramArray = new object[1];
			paramArray[0] = bstrQueryName;
			Invoker.Method(this, "Delete", paramArray);
		}

		#endregion

	}
}