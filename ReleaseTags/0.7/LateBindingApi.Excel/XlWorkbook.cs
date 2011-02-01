using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;

using System.Security;
using System.Security.Permissions;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Windows;
using LateBindingApi.Excel.VBIDE;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Charts;
using LateBindingApi.Excel.Office;
using LateBindingApi.Excel.Pivot;
using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Views; 
using LateBindingApi.Excel.Interfaces;

namespace LateBindingApi.Excel
{
    /// <summary>
    /// Represents a Workbook in Excel
    /// </summary>
    public class XlWorkbook : XlNonCreatable , IXlEventBinding 
    {
        #region Fields

        XlWorkbookEvents _eventBridge;

        #endregion

        #region Construction

        internal XlWorkbook(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion 

        #region COMReference Properties

        public LateBindingApi.Excel.Windows.XlWindows Windows
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Windows", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                LateBindingApi.Excel.Windows.XlWindows newClass = new LateBindingApi.Excel.Windows.XlWindows(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWebOptions WebOptions
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WebOptions", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWebOptions newClass = new XlWebOptions(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDocumentProperties BuiltinDocumentProperties
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("BuiltinDocumentProperties", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDocumentProperties newClass = new XlDocumentProperties(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlDocumentProperties CustomDocumentProperties
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CustomDocumentProperties", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDocumentProperties newClass = new XlDocumentProperties(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

 
        public XlCommandBars CommandBars 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CommandBars", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCommandBars newClass = new XlCommandBars(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlVBProject VBProject
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBProject", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlVBProject newClass = new XlVBProject(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        public XlStyles Styles
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Styles", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlStyles newClass = new XlStyles(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlSmartTagOptions SmartTagOptions
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SmartTagOptions", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlSmartTagOptions newClass = new XlSmartTagOptions(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlSheets Sheets
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Sheets", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlSheets newClass = new XlSheets(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlSheets Excel4IntlMacroSheets 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Excel4IntlMacroSheets", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlSheets newClass = new XlSheets(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlNonCreatable Parent
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Parent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlNonCreatable newClass = XlDynamicType.CreateDynamicType(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlSheets Excel4MacroSheets 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("Excel4MacroSheets", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlSheets newClass = new XlSheets(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlRoutingSlip RoutingSlip
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RoutingSlip", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlRoutingSlip newClass = new XlRoutingSlip(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlPublishObjects PublishObjects
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PublishObjects", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlPublishObjects newClass = new XlPublishObjects(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlNames Names
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Names", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlNames newClass = new XlNames(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlMailer Mailer
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Mailer", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlMailer newClass = new XlMailer(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCustomViews CustomViews
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CustomViews", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCustomViews newClass = new XlCustomViews(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlApplication Application
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Application", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlApplication newClass = new XlApplication(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWorksheets Worksheets
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Worksheets", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorksheets newClass = new XlWorksheets(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCharts ActiveChart
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveChart", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCharts newClass = new XlCharts(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlWorksheet ActiveSheet
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveSheet", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlWorksheet newClass = new XlWorksheet(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCharts Charts
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Charts", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlCharts newClass = new XlCharts(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Scalar Properties

        public string WriteReservedBy  
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WriteReservedBy", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public bool WriteReserved 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WriteReserved", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("WriteReserved", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool WritePassword 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WritePassword", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("WritePassword", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool VBASigned  
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("VBASigned", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool UpdateRemoteReferences 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UpdateRemoteReferences", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("UpdateRemoteReferences", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool UpdateLinks 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UpdateLinks", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("UpdateLinks", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool TemplateRemoveExtData 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TemplateRemoveExtData", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("TemplateRemoveExtData", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowPivotTableFieldList 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowPivotTableFieldList", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ShowPivotTableFieldList", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowConflictHistory 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowConflictHistory", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ShowConflictHistory", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
        
        public bool Routed 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Routed", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public int RevisionNumber  
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RevisionNumber", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public bool RemovePersonalInformation 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RemovePersonalInformation", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("RemovePersonalInformation", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ReadOnlyRecommended 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ReadOnlyRecommended", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ReadOnlyRecommended", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

       
        public bool ProtectWindows 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectWindows", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ProtectStructure  
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProtectStructure", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool PrecisionAsDisplayed 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PrecisionAsDisplayed", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("PrecisionAsDisplayed", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool PersonalViewPrintSettings 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PersonalViewPrintSettings", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("PersonalViewPrintSettings", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool PersonalViewListSettings 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PersonalViewListSettings", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("PersonalViewListSettings", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Path
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Path", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string PasswordEncryptionProvider 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PasswordEncryptionProvider", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public int PasswordEncryptionKeyLength 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PasswordEncryptionKeyLength", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public bool PasswordEncryptionFileProperties 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PasswordEncryptionFileProperties", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public string PasswordEncryptionAlgorithm  
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PasswordEncryptionAlgorithm", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string Password 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Password", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Password", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool MultiUserEditing 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MultiUserEditing", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool ListChangesOnNewSheet 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ListChangesOnNewSheet", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ListChangesOnNewSheet", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool KeepChangeHistory 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("KeepChangeHistory", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("KeepChangeHistory", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool IsInplace 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("IsInplace", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool IsAddin 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("IsAddin", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool HighlightChangesOnScreen 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HighlightChangesOnScreen", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HighlightChangesOnScreen", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasRoutingSlip 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasRoutingSlip", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("HasRoutingSlip", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool HasPassword 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("HasPassword", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public string FullNameURLEncoded
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FullNameURLEncoded", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }            
        }

        public XlFileFormat FileFormat 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FileFormat", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlFileFormat)returnValue;
            }
        }

        public bool EnvelopeVisible 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EnvelopeVisible", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("EnvelopeVisible", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool EnableAutoRecover 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EnableAutoRecover", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("EnableAutoRecover", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlDisplayDrawingObjects DisplayDrawingObjects 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayDrawingObjects", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlDisplayDrawingObjects)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("DisplayDrawingObjects", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Date1904
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Date1904", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }
       
        public XlCreator Creator 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public bool CreateBackup
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CreateBackup", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public XlSaveConflictResolution ConflictResolution 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ConflictResolution", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlSaveConflictResolution)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ConflictResolution", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int ChangeHistoryDuration 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ChangeHistoryDuration", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("ChangeHistoryDuration", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string CodeName 
        {
            get
            {
                object returnValue = InstanceType.InvokeMember("CodeName", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("CodeName", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int CalculationVersion  
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CalculationVersion", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public bool AutoUpdateSaveChanges 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoUpdateSaveChanges", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoUpdateSaveChanges", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int AutoUpdateFrequency 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoUpdateFrequency", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AutoUpdateFrequency", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AcceptLabelsInFormulas 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AcceptLabelsInFormulas", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("AcceptLabelsInFormulas", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Saved
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Saved", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] paramArray = new object[1];
                paramArray[0] = value;
                InstanceType.InvokeMember("Saved", BindingFlags.SetProperty, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ReadOnly
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ReadOnly", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }   
        }

        public string Name 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }             
        }

        public string FullName
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FullName", BindingFlags.GetProperty | BindingFlags.OptionalParamBinding, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }            
        }

        #endregion

        #region Methods
       
        public void AcceptAllChanges()
        {
            InstanceType.InvokeMember("AcceptAllChanges", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void OpenLinks(string name)
        {
            object[] paramArray = new object[1];
            paramArray[0] = name;
            InstanceType.InvokeMember("OpenLinks", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public LateBindingApi.Excel.Windows.XlWindows NewWindow()
        {
            object returnValue  = InstanceType.InvokeMember("NewWindow", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null;
            LateBindingApi.Excel.Windows.XlWindows newClass = new LateBindingApi.Excel.Windows.XlWindows(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }

        public void MergeWorkbook(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("MergeWorkbook", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public bool HighlightChangesOptions()
        {
            object returnValue  = InstanceType.InvokeMember("HighlightChangesOptions", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public bool ForwardMailer()
        {
            object returnValue  = InstanceType.InvokeMember("ForwardMailer", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public void FollowHyperlink(string address)
        {
            object[] paramArray = new object[1];
            paramArray[0] = address;
            InstanceType.InvokeMember("FollowHyperlink", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public bool ExclusiveAccess()
        {
            object returnValue  = InstanceType.InvokeMember("ExclusiveAccess", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public void EndReview()
        {
            InstanceType.InvokeMember("EndReview", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void DeleteNumberFormat(string numberFormat)
        {
            object[] paramArray = new object[1];
            paramArray[0] = numberFormat;
            InstanceType.InvokeMember("DeleteNumberFormat", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void LinkInfo(string name, XlLinkInfo linkInfo)
        {
            object[] paramArray = new object[3];
            paramArray[0] = name;
            paramArray[1] = linkInfo;
            InstanceType.InvokeMember("LinkInfo", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PrintOut()
        {
            InstanceType.InvokeMember("PrintOut", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PrintPreview()
        {
            InstanceType.InvokeMember("PrintPreview", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Protect()
        {
            InstanceType.InvokeMember("Protect", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }
 
        public void ProtectSharing()
        {
            InstanceType.InvokeMember("ProtectSharing", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void PurgeChangeHistoryNow(int days)
        {
            object[] paramArray = new object[1];
            paramArray[0] = days;
            InstanceType.InvokeMember("PurgeChangeHistoryNow", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RecheckSmartTags()
        {
            InstanceType.InvokeMember("RecheckSmartTags", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RefreshAll()
        {
            InstanceType.InvokeMember("RefreshAll", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RejectAllChanges()
        {
            InstanceType.InvokeMember("RejectAllChanges", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ReloadAs(MsoEncoding encoding)
        {
            object[] paramArray = new object[1];
            paramArray[0] = encoding;
            InstanceType.InvokeMember("ReloadAs", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RemoveUser(int index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            InstanceType.InvokeMember("RemoveUser", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Reply()
        {
            InstanceType.InvokeMember("Reply", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ReplyAll()
        {
            InstanceType.InvokeMember("ReplyAll", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ReplyWithChanges()
        {
            InstanceType.InvokeMember("ReplyWithChanges", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ResetColors()
        {
            InstanceType.InvokeMember("ResetColors", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Route()
        {
            InstanceType.InvokeMember("Route", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void RunAutoMacros(XlRunAutoMacro which)
        {
            object[] paramArray = new object[1];
            paramArray[0] = which;
            InstanceType.InvokeMember("RunAutoMacros", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SendForReview()
        {
            InstanceType.InvokeMember("SendForReview", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SendMail()
        {
            InstanceType.InvokeMember("SendMail", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SendMailer()
        {
            InstanceType.InvokeMember("SendMailer", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetLinkOnData(string name)
        {
            object[] paramArray = new object[1];
            paramArray[0] = name;
            InstanceType.InvokeMember("SetLinkOnData", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SetPasswordEncryptionOptions()
        {
            InstanceType.InvokeMember("SetPasswordEncryptionOptions", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Unprotect()
        {
            InstanceType.InvokeMember("Unprotect", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void UnprotectSharing()
        {
            InstanceType.InvokeMember("UnprotectSharing", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void UpdateFromFile()
        {
            InstanceType.InvokeMember("UpdateFromFile", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void UpdateLink()
        {
            InstanceType.InvokeMember("UpdateLink", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void WebPagePreview()
        {
            InstanceType.InvokeMember("WebPagePreview", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Close()
        {
            InstanceType.InvokeMember("Close", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public bool CheckIn()
        {
            object returnValue  = InstanceType.InvokeMember("CheckIn", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue;
        }

        public void ChangeLink(string name, string newName, XlLinkType type)
        {
            object[] paramArray = new object[3];
            paramArray[0] = name;
            paramArray[1] = newName;
            paramArray[2] = type;
            InstanceType.InvokeMember("ChangeLink", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ChangeFileAccess(XlFileAccess mode)
        {
            object[] paramArray = new object[1];
            paramArray[0] = mode;
            InstanceType.InvokeMember("ChangeFileAccess", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public bool CanCheckIn()
        {
            object returnValue  = InstanceType.InvokeMember("CanCheckIn", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
            return (bool)returnValue; 
        }

        public void BreakLink(string name, XlLinkType type)
        {
            object[] paramArray = new object[2];
            paramArray[0] = name;
            paramArray[1] = type;
            InstanceType.InvokeMember("BreakLink", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void AddToFavorites()
        {
            InstanceType.InvokeMember("AddToFavorites", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Activate()
        {
            InstanceType.InvokeMember("Activate", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Save()
        {
            InstanceType.InvokeMember("Save", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);                  
        }

        public void SaveAs(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void SaveCopyAs(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("SaveCopyAs", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
 
        }

        public void Close(bool saveChanges) 
        {
            object[] paramArray = new object[1];
            paramArray[0] = saveChanges;
            InstanceType.InvokeMember("Close", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);         
        }

        public void Close(string fileName)
        {
            object[] paramArray = new object[1];
            paramArray[0] = fileName;
            InstanceType.InvokeMember("Close", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture); 
        }
     
        #endregion

        #region Events

        /// <summary>
        /// Event Activate mapped to ActivateEvent. Reason: Workbook has an Activate Method
        /// </summary>
        public event WorkbookEvents_ActivateEventHandler ActivateEvent;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseActivateEvent()
        {
            if (null == ActivateEvent)
            {
                return;
            }

            ActivateEvent();
        }

        public event WorkbookEvents_AddinInstallEventHandler AddinInstall;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]

        internal void RaiseAddinInstallEvent()
        {
            if (null == AddinInstall)
            {
                return;
            }

            AddinInstall();
        }

        public event WorkbookEvents_AddinUninstallEventHandler AddinUninstall;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseAddinUninstallEvent()
        {
            if (null == AddinUninstall)
            {
                return;
            }

            AddinUninstall();
        }

        public event WorkbookEvents_BeforeCloseEventHandler BeforeClose;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseBeforeCloseEvent(ref bool Cancel)
        {
            if (null == BeforeClose)
            {
                return;
            }

            BeforeClose(ref Cancel);
        }

        public event WorkbookEvents_BeforePrintEventHandler BeforePrint;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseBeforePrintEvent(ref bool Cancel)
        {
            if (null == BeforePrint) return;

            BeforePrint(ref Cancel);
        }

        public event WorkbookEvents_BeforeSaveEventHandler BeforeSave;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseBeforeSaveEvent(bool SaveAsUI, ref bool Cancel)
        {
            if (null == BeforeSave) return;

            BeforeSave(SaveAsUI, ref Cancel);
        }

        public event WorkbookEvents_DeactivateEventHandler Deactivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseDeactivateEvent()
        {
            if (null == Deactivate) return;

            Deactivate();
        }

        public event WorkbookEvents_NewSheetEventHandler NewSheet;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseNewSheetEvent(object Sh)
        {
            if (null == NewSheet)
            {
                Marshal.ReleaseComObject(Sh);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            NewSheet(ws);
        }

        public event WorkbookEvents_OpenEventHandler Open;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseOpenEvent()
        {
            if (null == Open) return;

            Open();
        }

        public event WorkbookEvents_PivotTableCloseConnectionEventHandler PivotTableCloseConnection;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaisePivotTableCloseConnectionEvent(object Target)
        {
            if (null == PivotTableCloseConnection)
            {
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlPivotTable pt = new XlPivotTable(this, Target);
            ListChildReferences.Add(pt);

            PivotTableCloseConnection(pt);
        }

        public event WorkbookEvents_PivotTableOpenConnectionEventHandler PivotTableOpenConnection;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaisePivotTableOpenConnectionEvent(object Target)
        {
            if (null == PivotTableOpenConnection)
            {
                Marshal.ReleaseComObject(Target);
                return;
            } 

            XlPivotTable pt = new XlPivotTable(this, Target);
            ListChildReferences.Add(pt);

            PivotTableOpenConnection(pt);
        }

        public event WorkbookEvents_SheetActivateEventHandler SheetActivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetActivateEvent(object Sh)
        {
            if (null == SheetActivate)
            {
                Marshal.ReleaseComObject(Sh);
                return;
            }
            
            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            SheetActivate(ws);
        }

        public event WorkbookEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetBeforeDoubleClickEvent(object Sh, object Target,ref bool Cancel)
        {
            if (null == SheetBeforeDoubleClick)
            {
                Marshal.ReleaseComObject(Sh);
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);
            
            XlRange ra = new XlRange(this, Sh);
            ListChildReferences.Add(ra);

            SheetBeforeDoubleClick(ws, ra, ref Cancel);
        }

        public event WorkbookEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClick;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetBeforeRightClickEvent(object Sh, object Target, ref bool Cancel)
        {
            if (null == SheetBeforeRightClick)
            {
                Marshal.ReleaseComObject(Sh);
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            XlRange ra = new XlRange(this, Sh);
            ListChildReferences.Add(ra);

            SheetBeforeRightClick(ws, ra, ref Cancel);
        }

        public event WorkbookEvents_SheetCalculateEventHandler SheetCalculate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetCalculateEvent(object Sh)
        {
            if (null == SheetCalculate)
            {
                Marshal.ReleaseComObject(Sh);            
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            SheetCalculate(ws);
        }

        public event WorkbookEvents_SheetChangeEventHandler SheetChange;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetChangeEvent(object Sh, object Target)
        {
            if (null == SheetChange)
            {
                Marshal.ReleaseComObject(Sh);
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            XlRange ra = new XlRange(this, Sh);
            ListChildReferences.Add(ra);

            SheetChange(ws, ra);
        }

        public event WorkbookEvents_SheetDeactivateEventHandler SheetDeactivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetDeactivateEvent(object Sh)
        {
            if (null == SheetDeactivate)
            {
                Marshal.ReleaseComObject(Sh);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            SheetDeactivate(ws);
        }

        public event WorkbookEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlink;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetFollowHyperlinkEvent(object Sh, object Target)
        {
            if (null == SheetFollowHyperlink)
            {
                Marshal.ReleaseComObject(Sh);
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            XlHyperlink hl = new XlHyperlink(this, Target);
            ListChildReferences.Add(hl);

            SheetFollowHyperlink(ws, hl);
        }

        public event WorkbookEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetPivotTableUpdateEvent(object Sh, object Target)
        {
            if (null == SheetPivotTableUpdate)
            {
                Marshal.ReleaseComObject(Sh);
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            XlPivotTable pt = new XlPivotTable(this, Target);
            ListChildReferences.Add(pt);

            SheetPivotTableUpdate(ws, pt);
        }

        public event WorkbookEvents_SheetSelectionChangeEventHandler SheetSelectionChange;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetSelectionChangeEvent(object Sh, object Target)
        {
            if (null == SheetSelectionChange)
            {
                Marshal.ReleaseComObject(Sh);
                Marshal.ReleaseComObject(Target);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, Sh);
            ListChildReferences.Add(ws);

            XlRange ra = new XlRange(this, Sh);
            ListChildReferences.Add(ra);

            SheetSelectionChange(ws, ra);
        }

        public event WorkbookEvents_WindowActivateEventHandler WindowActivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWindowActivateEvent(object Wn)
        {
            if (null == WindowActivate)
            {
                Marshal.ReleaseComObject(Wn);
                return;
            }

             LateBindingApi.Excel.Windows.XlWindow win = new LateBindingApi.Excel.Windows.XlWindow(this, Wn);
             ListChildReferences.Add(win);

            WindowActivate(win);
        }

        public event WorkbookEvents_WindowDeactivateEventHandler WindowDeactivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWindowDeactivateEvent(object Wn)
        {
            if (null == WindowDeactivate)
            {
                Marshal.ReleaseComObject(Wn);
                return;
            }

            LateBindingApi.Excel.Windows.XlWindow win = new LateBindingApi.Excel.Windows.XlWindow(this, Wn);
            ListChildReferences.Add(win);

            WindowDeactivate(win);
        }

        public event WorkbookEvents_WindowResizeEventHandler WindowResize;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWindowResizeEvent(object Wn)
        {

            if (null == WindowResize)
            {
                Marshal.ReleaseComObject(Wn);
                return;
            }

            LateBindingApi.Excel.Windows.XlWindow win = new LateBindingApi.Excel.Windows.XlWindow(this, Wn);
            ListChildReferences.Add(win);

            WindowResize(win);
        }

        #endregion

        #region IXlEventBinding Members

        public void SetupEventBinding()
        {
            _eventBridge = new XlWorkbookEvents();
            _eventBridge.SetupEventBinding(this);
        }

        public void RemoveEventBinding()
        {
            _eventBridge.RemoveEventBinding();
        }

        #endregion
    }
}
