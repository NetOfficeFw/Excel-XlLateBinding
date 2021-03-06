//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
using System.Collections;
namespace LateBindingApi.Office
{
	[SupportByLibrary("OF12","OF14")]
	public class TextRange2 : _IMsoDispObj
	{
		#region Construction

		public TextRange2(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public TextRange2(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public TextRange2(COMObject replacedObject) : base(replacedObject)
		{
		}

		public TextRange2()
		{
		}

		#endregion

		#region Properties

		[SupportByLibrary("OF12","OF14")]
		public Int32 Count
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Count");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public IEnumerator GetEnumerator()
		{
			object enumProxy = Invoker.PropertyGet(this, "_NewEnum");
			COMObject enumerator = new COMObject(this, enumProxy);
			Invoker.Method(enumerator, "Reset", null);
			bool isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
            while (true == isMoveNextTrue)
            {
                object itemProxy = Invoker.PropertyGet(enumerator, "Current", null);
				LateBindingApi.Office.TextRange2 returnClass = new LateBindingApi.Office.TextRange2 (this, itemProxy);
				isMoveNextTrue = (bool)Invoker.MethodReturn(enumerator, "MoveNext", null);
				yield return returnClass;
            }
		}

		[SupportByLibrary("OF12","OF14")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 get_Paragraphs(Int32 start, Int32 length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "Paragraphs", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 get_Sentences(Int32 start, Int32 length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "Sentences", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 get_Words(Int32 start, Int32 length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "Words", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 get_Characters(Int32 start, Int32 length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "Characters", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 get_Lines(Int32 start, Int32 length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "Lines", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 get_Runs(Int32 start, Int32 length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "Runs", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.ParagraphFormat2 ParagraphFormat
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ParagraphFormat");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.ParagraphFormat2 newClass = new LateBindingApi.Office.ParagraphFormat2(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Font2 Font
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Font");
				if(null == returnValue)
					return null;
				LateBindingApi.Office.Font2 newClass = new LateBindingApi.Office.Font2(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Int32 Length
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Length");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Int32 Start
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Start");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Double BoundLeft
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BoundLeft");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Double BoundTop
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BoundTop");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Double BoundWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BoundWidth");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public Double BoundHeight
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BoundHeight");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.Enums.MsoLanguageID LanguageID
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LanguageID");
				return (LateBindingApi.Office.Enums.MsoLanguageID)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LanguageID", value);
			}						
		}


		[SupportByLibrary("OF14")]
		public LateBindingApi.Office.TextRange2 get_MathZones(Int32 start, Int32 length)
		{
			object[] paramArray = new object[2];
			paramArray[0] = start;
			paramArray[1] = length;
			object returnValue = Invoker.PropertyGet(this, "MathZones", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		#endregion

		#region Methods

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 this[object index]
		{
			get
			{
				object[] paramArray = new object[1];
				paramArray[0] = index;		
				object returnValue = Invoker.MethodReturn(this, "Item", paramArray);
				if(null == returnValue)
					return null;
				LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 TrimText()
		{
			object returnValue = Invoker.MethodReturn(this, "TrimText", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 InsertAfter(string newText)
		{
			object[] paramArray = new object[1];
			paramArray[0] = newText;
			object returnValue = Invoker.MethodReturn(this, "InsertAfter", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 InsertBefore(string newText)
		{
			object[] paramArray = new object[1];
			paramArray[0] = newText;
			object returnValue = Invoker.MethodReturn(this, "InsertBefore", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 InsertSymbol(string fontName, Int32 charNumber, LateBindingApi.Office.Enums.MsoTriState unicode)
		{
			object[] paramArray = new object[3];
			paramArray[0] = fontName;
			paramArray[1] = charNumber;
			paramArray[2] = unicode;
			object returnValue = Invoker.MethodReturn(this, "InsertSymbol", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public void Select()
		{
			Invoker.Method(this, "Select", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void Cut()
		{
			Invoker.Method(this, "Cut", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void Copy()
		{
			Invoker.Method(this, "Copy", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void Delete()
		{
			Invoker.Method(this, "Delete", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 Paste()
		{
			object returnValue = Invoker.MethodReturn(this, "Paste", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 PasteSpecial(LateBindingApi.Office.Enums.MsoClipboardFormat format)
		{
			object[] paramArray = new object[1];
			paramArray[0] = format;
			object returnValue = Invoker.MethodReturn(this, "PasteSpecial", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public void ChangeCase(LateBindingApi.Office.Enums.MsoTextChangeCase type)
		{
			object[] paramArray = new object[1];
			paramArray[0] = type;
			Invoker.Method(this, "ChangeCase", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void AddPeriods()
		{
			Invoker.Method(this, "AddPeriods", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void RemovePeriods()
		{
			Invoker.Method(this, "RemovePeriods", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 Find(string findWhat, Int32 after, LateBindingApi.Office.Enums.MsoTriState matchCase, LateBindingApi.Office.Enums.MsoTriState wholeWords)
		{
			object[] paramArray = new object[4];
			paramArray[0] = findWhat;
			paramArray[1] = after;
			paramArray[2] = matchCase;
			paramArray[3] = wholeWords;
			object returnValue = Invoker.MethodReturn(this, "Find", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public LateBindingApi.Office.TextRange2 Replace(string findWhat, string replaceWhat, Int32 after, LateBindingApi.Office.Enums.MsoTriState matchCase, LateBindingApi.Office.Enums.MsoTriState wholeWords)
		{
			object[] paramArray = new object[5];
			paramArray[0] = findWhat;
			paramArray[1] = replaceWhat;
			paramArray[2] = after;
			paramArray[3] = matchCase;
			paramArray[4] = wholeWords;
			object returnValue = Invoker.MethodReturn(this, "Replace", paramArray);
			if(null == returnValue)
				return null;
			LateBindingApi.Office.TextRange2 newClass = new LateBindingApi.Office.TextRange2(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("OF12","OF14")]
		public void RotatedBounds(Double x1, Double y1, Double x2, Double y2, Double x3, Double y3, Double x4, Double y4)
		{
			object[] paramArray = new object[8];
			paramArray.SetValue(x1,0);
			paramArray.SetValue(y1,1);
			paramArray.SetValue(x2,2);
			paramArray.SetValue(y2,3);
			paramArray.SetValue(x3,4);
			paramArray.SetValue(y3,5);
			paramArray.SetValue(x4,6);
			paramArray.SetValue(y4,7);
			Invoker.Method(this, "RotatedBounds", paramArray);
		}

		[SupportByLibrary("OF12","OF14")]
		public void RtlRun()
		{
			Invoker.Method(this, "RtlRun", null);
		}

		[SupportByLibrary("OF12","OF14")]
		public void LtrRun()
		{
			Invoker.Method(this, "LtrRun", null);
		}

		#endregion

	}
}
