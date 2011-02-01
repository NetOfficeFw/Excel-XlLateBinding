//Generated by LateBindingApi.CodeGenerator
using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using LateBindingApi.Core;
namespace LateBindingApi.Excel
{
	[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
	public class PageSetup : COMObject
	{
		#region Construction

		public PageSetup(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}

		public PageSetup(COMObject parentObject, object comProxy, Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		public PageSetup(COMObject replacedObject) : base(replacedObject)
		{
		}

		public PageSetup()
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
		public bool BlackAndWhite
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BlackAndWhite");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "BlackAndWhite", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double BottomMargin
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "BottomMargin");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "BottomMargin", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string CenterFooter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CenterFooter");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CenterFooter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string CenterHeader
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CenterHeader");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CenterHeader", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool CenterHorizontally
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CenterHorizontally");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CenterHorizontally", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool CenterVertically
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CenterVertically");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "CenterVertically", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlObjectSize ChartSize
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ChartSize");
				return (LateBindingApi.Excel.Enums.XlObjectSize)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ChartSize", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Draft
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Draft");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Draft", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 FirstPageNumber
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FirstPageNumber");
				return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FirstPageNumber", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Int32 FitToPagesTall
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FitToPagesTall");
                return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FitToPagesTall", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
        public Int32 FitToPagesWide
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FitToPagesWide");
                return (Int32)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FitToPagesWide", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double FooterMargin
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FooterMargin");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "FooterMargin", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double HeaderMargin
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "HeaderMargin");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "HeaderMargin", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string LeftFooter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LeftFooter");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LeftFooter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string LeftHeader
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LeftHeader");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LeftHeader", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double LeftMargin
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LeftMargin");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "LeftMargin", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlOrder Order
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Order");
				return (LateBindingApi.Excel.Enums.XlOrder)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Order", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlPageOrientation Orientation
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Orientation");
				return (LateBindingApi.Excel.Enums.XlPageOrientation)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "Orientation", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlPaperSize PaperSize
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PaperSize");
				return (LateBindingApi.Excel.Enums.XlPaperSize)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PaperSize", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string PrintArea
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintArea");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintArea", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool PrintGridlines
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintGridlines");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintGridlines", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool PrintHeadings
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintHeadings");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintHeadings", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool PrintNotes
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintNotes");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintNotes", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant PrintQuality
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintQuality", null);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
			}
			set
			{
				Invoker.PropertySet(this, "PrintQuality", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public COMVariant get_PrintQuality(object index)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
				object returnValue = Invoker.PropertyGet(this, "PrintQuality", paramArray);
				COMVariant returnObject = LateBindingApi.Core.Factory.CreateVariantFromComProxy(this, returnValue);
				return returnObject;
		}
		
		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public void set_PrintQuality(object index, object value)
		{
			object[] paramArray = new object[1];
			paramArray[0] = index;
			Invoker.PropertySet(this, "PrintQuality", paramArray, value);
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string PrintTitleColumns
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintTitleColumns");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintTitleColumns", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string PrintTitleRows
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintTitleRows");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintTitleRows", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string RightFooter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RightFooter");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RightFooter", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public string RightHeader
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RightHeader");
				return (string)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RightHeader", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double RightMargin
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RightMargin");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "RightMargin", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public Double TopMargin
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "TopMargin");
				return (Double)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "TopMargin", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public bool Zoom
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Zoom");
                return (bool)returnValue; 
			}
			set
			{
				Invoker.PropertySet(this, "Zoom", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14","XL9")]
		public LateBindingApi.Excel.Enums.XlPrintLocation PrintComments
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintComments");
				return (LateBindingApi.Excel.Enums.XlPrintLocation)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintComments", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Enums.XlPrintErrors PrintErrors
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "PrintErrors");
				return (LateBindingApi.Excel.Enums.XlPrintErrors)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "PrintErrors", value);
			}						
		}


		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Graphic CenterHeaderPicture
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CenterHeaderPicture");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Graphic newClass = new LateBindingApi.Excel.Graphic(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Graphic CenterFooterPicture
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "CenterFooterPicture");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Graphic newClass = new LateBindingApi.Excel.Graphic(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Graphic LeftHeaderPicture
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LeftHeaderPicture");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Graphic newClass = new LateBindingApi.Excel.Graphic(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Graphic LeftFooterPicture
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "LeftFooterPicture");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Graphic newClass = new LateBindingApi.Excel.Graphic(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Graphic RightHeaderPicture
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RightHeaderPicture");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Graphic newClass = new LateBindingApi.Excel.Graphic(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL10","XL11","XL12","XL14")]
		public LateBindingApi.Excel.Graphic RightFooterPicture
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "RightFooterPicture");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Graphic newClass = new LateBindingApi.Excel.Graphic(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public bool OddAndEvenPagesHeaderFooter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "OddAndEvenPagesHeaderFooter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "OddAndEvenPagesHeaderFooter", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public bool DifferentFirstPageHeaderFooter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "DifferentFirstPageHeaderFooter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "DifferentFirstPageHeaderFooter", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public bool ScaleWithDocHeaderFooter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "ScaleWithDocHeaderFooter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "ScaleWithDocHeaderFooter", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public bool AlignMarginsHeaderFooter
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "AlignMarginsHeaderFooter");
				return (bool)returnValue;
			}
			set
			{
				Invoker.PropertySet(this, "AlignMarginsHeaderFooter", value);
			}						
		}


		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Pages Pages
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "Pages");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Pages newClass = new LateBindingApi.Excel.Pages(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Page EvenPage
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "EvenPage");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Page newClass = new LateBindingApi.Excel.Page(this, returnValue);
				return newClass;
			}
		}

		[SupportByLibrary("XL12","XL14")]
		public LateBindingApi.Excel.Page FirstPage
		{
			get
			{
				object returnValue = Invoker.PropertyGet(this, "FirstPage");
				if(null == returnValue)
					return null;
				LateBindingApi.Excel.Page newClass = new LateBindingApi.Excel.Page(this, returnValue);
				return newClass;
			}
		}

		#endregion

		#region Methods

		#endregion

	}
}
