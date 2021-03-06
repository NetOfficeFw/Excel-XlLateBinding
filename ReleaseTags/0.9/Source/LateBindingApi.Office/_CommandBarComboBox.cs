//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
	public class _CommandBarComboBox : CommandBarControl
	{
		#region Construction

		public _CommandBarComboBox(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public _CommandBarComboBox(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public _CommandBarComboBox(COMObject replacedObject) : base(replacedObject)
		{
		}

		public _CommandBarComboBox()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 DropDownLines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DropDownLines");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DropDownLines", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 DropDownWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DropDownWidth");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DropDownWidth", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public string get_List(Int32 index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
				object returnValue = Invoker.PropertyGet(this, "List", paramArray);
				return (string)returnValue;
		}
		
		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void set_List(Int32 index, object value)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.PropertySet(this, "List", paramArray, value);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 ListCount
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListCount");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 ListHeaderCount
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListHeaderCount");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ListHeaderCount", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public Int32 ListIndex
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ListIndex");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ListIndex", value);
			}						
		}


		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public LateBindingApi.Office.Enums.MsoComboStyle Style
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Style");
				return (LateBindingApi.Office.Enums.MsoComboStyle)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Style", value);
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


		[SupportByLibrary("OF14")]
		public COMVariant InstanceIdPtr
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "InstanceIdPtr");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void AddItem(string text)
		{
			object[] paramArray = new object[1];
			paramArray[0] = text;
			Invoker.Method(this, "AddItem", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void AddItem(string text, object index)
		{
			object[] paramArray = new object[2];
			paramArray[0] = text;
			paramArray[1] = index;
			Invoker.Method(this, "AddItem", paramArray);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void Clear()
		{
			Invoker.Method(this, "Clear", null);
		}

		[SupportByLibrary("OF10","OF11","OF12","OF14","OF9")]
		public void RemoveItem(Int32 index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.Method(this, "RemoveItem", paramArray);
		}

		#endregion

	}
}
