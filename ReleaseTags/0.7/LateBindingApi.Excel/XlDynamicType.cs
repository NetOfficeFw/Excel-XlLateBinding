using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel
{
    internal static class XlDynamicType
    {
        public class ProxyTypeException : Exception
        {
            public ProxyTypeException(string text): base(text)
            { }
        }

        internal static XlNonCreatable CreateDynamicType(IXlObject parent, object comProxy)
        {
            string className = TypeDescriptor.GetClassName(comProxy);
            switch (className)
            {               

                case "Workbooks":

                    XlWorkbooks newBooks = new XlWorkbooks(parent, comProxy);
                    return newBooks;

                case "Workbook":

                    XlWorkbook newBook = new XlWorkbook(parent, comProxy);
                    return newBook;

                case "Worksheets":

                    XlWorksheets newSheets = new XlWorksheets(parent, comProxy);
                    return newSheets;

                case "Worksheet":

                    XlWorksheet newSheet = new XlWorksheet(parent, comProxy);
                    return newSheet;

                case "Range":

                    XlRange newRange = new XlRange(parent, comProxy);
                    return newRange;

                default:

                    throw (new ProxyTypeException("Unhandled ComProxyType: " + className));

            }
        }
    }
}
