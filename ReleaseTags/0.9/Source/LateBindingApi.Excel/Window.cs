//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Application Application
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Application");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Application newClass = new LateBindingApi.Excel.Application(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlCreator Creator
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Creator");
				return (LateBindingApi.Excel.Enums.XlCreator)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Parent
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Parent");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range ActiveCell
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ActiveCell");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Chart ActiveChart
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ActiveChart");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Chart newClass = new LateBindingApi.Excel.Chart(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Pane ActivePane
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ActivePane");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Pane newClass = new LateBindingApi.Excel.Pane(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject ActiveSheet
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ActiveSheet");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Caption
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Caption");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Caption", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayFormulas
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayFormulas");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayFormulas", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayGridlines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayGridlines");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayGridlines", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayHeadings
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayHeadings");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayHeadings", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayHorizontalScrollBar
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayHorizontalScrollBar");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayHorizontalScrollBar", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayOutline
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayOutline");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayOutline", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool _DisplayRightToLeft
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "_DisplayRightToLeft");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "_DisplayRightToLeft", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayVerticalScrollBar
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayVerticalScrollBar");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayVerticalScrollBar", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayWorkbookTabs
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayWorkbookTabs");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayWorkbookTabs", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayZeros
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayZeros");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayZeros", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool EnableResize
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "EnableResize");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "EnableResize", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool FreezePanes
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FreezePanes");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FreezePanes", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 GridlineColor
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "GridlineColor");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "GridlineColor", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlColorIndex GridlineColorIndex
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "GridlineColorIndex");
				return (LateBindingApi.Excel.Enums.XlColorIndex)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "GridlineColorIndex", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Height
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Height");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Height", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 Index
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Index");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Left
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Left");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Left", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string OnWindow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OnWindow");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OnWindow", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Panes Panes
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Panes");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Panes newClass = new LateBindingApi.Excel.Panes(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range RangeSelection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RangeSelection");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ScrollColumn
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ScrollColumn");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ScrollColumn", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 ScrollRow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ScrollRow");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ScrollRow", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Sheets SelectedSheets
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SelectedSheets");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Sheets newClass = new LateBindingApi.Excel.Sheets(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject Selection
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Selection");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Split
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Split");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Split", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 SplitColumn
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SplitColumn");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SplitColumn", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double SplitHorizontal
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SplitHorizontal");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SplitHorizontal", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 SplitRow
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SplitRow");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SplitRow", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double SplitVertical
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SplitVertical");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "SplitVertical", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double TabRatio
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TabRatio");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TabRatio", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Top
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Top");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Top", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlWindowType Type
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Type");
				return (LateBindingApi.Excel.Enums.XlWindowType)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double UsableHeight
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "UsableHeight");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double UsableWidth
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "UsableWidth");
				return (Double)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
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


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Range VisibleRange
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "VisibleRange");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Range newClass = new LateBindingApi.Excel.Range(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double Width
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Width");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Width", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 WindowNumber
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WindowNumber");
				return (Int32)returnValue;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlWindowState WindowState
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "WindowState");
				return (LateBindingApi.Excel.Enums.XlWindowState)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "WindowState", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Zoom
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Zoom");
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "Zoom", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlWindowView View
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "View");
				return (LateBindingApi.Excel.Enums.XlWindowView)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "View", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool DisplayRightToLeft
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayRightToLeft");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayRightToLeft", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.SheetViews SheetViews
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "SheetViews");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.SheetViews newClass = new LateBindingApi.Excel.SheetViews(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public COMObject ActiveSheetView
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ActiveSheetView");
				COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
				return returnObject;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public bool DisplayRuler
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayRuler");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayRuler", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public bool AutoFilterDateGrouping
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AutoFilterDateGrouping");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AutoFilterDateGrouping", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public bool DisplayWhitespace
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DisplayWhitespace");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DisplayWhitespace", value);
			}						
		}


		#endregion

		#region Methods

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant Activate()
		{
			object returnValue = Invoker.MethodReturn(this, "Activate", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ActivateNext()
		{
			object returnValue = Invoker.MethodReturn(this, "ActivateNext", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ActivatePrevious()
		{
			object returnValue = Invoker.MethodReturn(this, "ActivatePrevious", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Close()
		{
			object returnValue = Invoker.MethodReturn(this, "Close", null);
			return (bool)returnValue;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Close(object saveChanges, object filename, object routeWorkbook)
		{
			object[] paramArray = new object[3];
			paramArray[0] = saveChanges;
			paramArray[1] = filename;
			paramArray[2] = routeWorkbook;
			object returnValue = Invoker.MethodReturn(this, "Close", paramArray);
			return (bool)returnValue;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant LargeScroll()
		{
			object returnValue = Invoker.MethodReturn(this, "LargeScroll", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant LargeScroll(object down, object up, object toRight, object toLeft)
		{
			object[] paramArray = new object[4];
			paramArray[0] = down;
			paramArray[1] = up;
			paramArray[2] = toRight;
			paramArray[3] = toLeft;
			object returnValue = Invoker.MethodReturn(this, "LargeScroll", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Window NewWindow()
		{
			object returnValue = Invoker.MethodReturn(this, "NewWindow", null);
			if(null == returnValue)
				return null;
			LateBindingApi.Excel.Window newClass = new LateBindingApi.Excel.Window(this, returnValue);
			return newClass;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant PrintOut()
		{
			object returnValue = Invoker.MethodReturn(this, "PrintOut", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			object[] paramArray = new object[8];
			paramArray[0] = from;
			paramArray[1] = to;
			paramArray[2] = copies;
			paramArray[3] = preview;
			paramArray[4] = activePrinter;
			paramArray[5] = printToFile;
			paramArray[6] = collate;
			paramArray[7] = prToFileName;
			object returnValue = Invoker.MethodReturn(this, "PrintOut", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant PrintPreview()
		{
			object returnValue = Invoker.MethodReturn(this, "PrintPreview", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant PrintPreview(object enableChanges)
		{
			object[] paramArray = new object[1];
			paramArray[0] = enableChanges;
			object returnValue = Invoker.MethodReturn(this, "PrintPreview", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ScrollWorkbookTabs()
		{
			object returnValue = Invoker.MethodReturn(this, "ScrollWorkbookTabs", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant ScrollWorkbookTabs(object sheets, object position)
		{
			object[] paramArray = new object[2];
			paramArray[0] = sheets;
			paramArray[1] = position;
			object returnValue = Invoker.MethodReturn(this, "ScrollWorkbookTabs", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SmallScroll()
		{
			object returnValue = Invoker.MethodReturn(this, "SmallScroll", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant SmallScroll(object down, object up, object toRight, object toLeft)
		{
			object[] paramArray = new object[4];
			paramArray[0] = down;
			paramArray[1] = up;
			paramArray[2] = toRight;
			paramArray[3] = toLeft;
			object returnValue = Invoker.MethodReturn(this, "SmallScroll", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 PointsToScreenPixelsX(Int32 points)
		{
			object[] paramArray = new object[1];
			paramArray[0] = points;
			object returnValue = Invoker.MethodReturn(this, "PointsToScreenPixelsX", paramArray);
			return (Int32)returnValue;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 PointsToScreenPixelsY(Int32 points)
		{
			object[] paramArray = new object[1];
			paramArray[0] = points;
			object returnValue = Invoker.MethodReturn(this, "PointsToScreenPixelsY", paramArray);
			return (Int32)returnValue;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMObject RangeFromPoint(Int32 x, Int32 y)
		{
			object[] paramArray = new object[2];
			paramArray[0] = x;
			paramArray[1] = y;
			object returnValue = Invoker.MethodReturn(this, "RangeFromPoint", paramArray);
			COMObject returnObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void ScrollIntoView(Int32 left, Int32 top, Int32 width, Int32 height)
		{
			object[] paramArray = new object[4];
			paramArray[0] = left;
			paramArray[1] = top;
			paramArray[2] = width;
			paramArray[3] = height;
			Invoker.Method(this, "ScrollIntoView", paramArray);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void ScrollIntoView(Int32 left, Int32 top, Int32 width, Int32 height, object start)
		{
			object[] paramArray = new object[5];
			paramArray[0] = left;
			paramArray[1] = top;
			paramArray[2] = width;
			paramArray[3] = height;
			paramArray[4] = start;
			Invoker.Method(this, "ScrollIntoView", paramArray);
		}

		[SupportByLibrary("XL12","XL14")]
		public COMVariant _PrintOut()
		{
			object returnValue = Invoker.MethodReturn(this, "_PrintOut", null);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		[SupportByLibrary("XL12","XL14")]
		public COMVariant _PrintOut(object from, object to, object copies, object preview, object activePrinter, object printToFile, object collate, object prToFileName)
		{
			object[] paramArray = new object[8];
			paramArray[0] = from;
			paramArray[1] = to;
			paramArray[2] = copies;
			paramArray[3] = preview;
			paramArray[4] = activePrinter;
			paramArray[5] = printToFile;
			paramArray[6] = collate;
			paramArray[7] = prToFileName;
			object returnValue = Invoker.MethodReturn(this, "_PrintOut", paramArray);
			COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
			return returnObject;
		}

		#endregion

	}
}
