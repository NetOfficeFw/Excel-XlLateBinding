using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Globalization;
using System.IO;
using System.ComponentModel;
using System.Security;
using System.Security.Permissions;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;
using LateBindingApi.Excel.Charts;
using LateBindingApi.Excel.Dialogs;
using LateBindingApi.Excel.Windows;
using LateBindingApi.Excel.Pivot;
using LateBindingApi.Excel.FileDialog;
using LateBindingApi.Excel.Styles;
using LateBindingApi.Excel.Web;
using LateBindingApi.Excel.Error;
using LateBindingApi.Excel.Odbc;
using LateBindingApi.Excel.OleDb;
using LateBindingApi.Excel.Recent;
using LateBindingApi.Excel.COMAddins;
using LateBindingApi.Excel.Addins;
using LateBindingApi.Excel.Recover;
using LateBindingApi.Excel.Language;
using LateBindingApi.Excel.Office;

[assembly: CLSCompliant(true)]
namespace LateBindingApi.Excel
{
    /// <summary>
    /// Represents the Excel Application, not bind to a specific Version
    /// </summary>
    public class XlApplication : XlCreatable, IXlEventBinding
    {   
        #region Constants

        // Default Progid from Excel without Version
        private const string _DefaultProgId = "Excel.Application";

        #endregion

        #region Fields

        XlApplicationEvents _eventBridge;

        #endregion       
	    
        #region Construction
		
        /// <summary>
        /// IXlCreatable Constructor
        /// </summary>
		public XlApplication()
		{         
            this.CreateCOMReference(_DefaultProgId);
        }

        public XlApplication(string progId)
        {
            this.CreateCOMReference(progId);
        }

        public XlApplication(IXlObject parentReference, object comReference) : base(parentReference, comReference)
		{

		}

		#endregion

        #region Scalar Properties

        public string ActivePrinter
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActivePrinter", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ActivePrinter", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AlertBeforeOverwriting
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AlertBeforeOverwriting", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AlertBeforeOverwriting", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AskToUpdateLinks
        {
            get
            {

                object returnValue  = InstanceType.InvokeMember("AskToUpdateLinks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
            set
            {
   
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AskToUpdateLinks", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AutoFormatAsYouTypeReplaceHyperlinks
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoFormatAsYouTypeReplaceHyperlinks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AutoFormatAsYouTypeReplaceHyperlinks", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool AutoPercentEntry
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoPercentEntry", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AutoPercentEntry", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }        
        }

        public XlMsoAutomationSecurity AutomationSecurity
        {
            get
            {                
                object returnValue  = InstanceType.InvokeMember("AutomationSecurity", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlMsoAutomationSecurity)returnValue;
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AutomationSecurity", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string AltStartupPath
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AltStartupPath", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("AltStartupPath", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Build
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Build", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue; 
            }
        }
 
        public bool CalculateBeforeSave
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CalculateBeforeSave", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("CalculateBeforeSave", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }

        }
        
        public XlCalculation Calculation
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Calculation", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCalculation)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Calculation", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlCalculationInterruptKey CalculationInterruptKey
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CalculationInterruptKey", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCalculationInterruptKey)returnValue;
            }
            set
            {         
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("CalculationInterruptKey", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlCalculationState CalculationState 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CalculationState", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCalculationState)returnValue;
            }
        }

        public int CalculationVersion 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CalculationVersion", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue; 
            }
        }

        public bool CanPlaySounds
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CanPlaySounds", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
        }

        public bool CanRecordSounds
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CanRecordSounds", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
        }

        public string Caption
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Caption", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue; 
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Caption", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool CellDragAndDrop
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CellDragAndDrop", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("CellDragAndDrop", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);

            }
        }

        public bool ConstrainNumeric
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ConstrainNumeric", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ConstrainNumeric", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);

            }
        }

        public double ControlCharacters
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ControlCharacters", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ControlCharacters", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);

            }
        }

        public bool CopyObjectsWithCells
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CopyObjectsWithCells", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("CopyObjectsWithCells", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);

            }
        }

        public XlCreator Creator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Creator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCreator)returnValue;
            }
        }

        public double CursorMovement
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CursorMovement", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("CursorMovement", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double CustomListCount
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CustomListCount", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
        }

        public XlCutCopyMode CutCopyMode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CutCopyMode", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCutCopyMode)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("CutCopyMode", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int DataEntryMode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DataEntryMode", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DataEntryMode", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string DefaultFilePath
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DefaultFilePath", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DefaultFilePath", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int DefaultSheetDirection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DefaultSheetDirection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DefaultSheetDirection", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }
 
        public double DDEAppReturnCode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DDEAppReturnCode", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
        }

        public string Version
        {         
            get
            {
                object returnValue  = InstanceType.InvokeMember("Version", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue; 
            }
        }

        public string DecimalSeparator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DecimalSeparator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue; 
            }
            set
            {               
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DecimalSeparator", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayAlerts
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayAlerts", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayClipboardWindow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayClipboardWindow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayClipboardWindow", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlCommentDisplayMode DisplayCommentIndicator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayCommentIndicator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlCommentDisplayMode)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayCommentIndicator", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayExcel4Menus
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayExcel4Menus", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayExcel4Menus", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayFormulaBar
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayFormulaBar", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayFormulaBar", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayFullScreen
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayFullScreen", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayFullScreen", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayFunctionToolTips
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayFunctionToolTips", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayFunctionToolTips", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayInsertOptions
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayInsertOptions", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayInsertOptions", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayNoteIndicator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayNoteIndicator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayNoteIndicator", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayPasteOptions
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayPasteOptions", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayPasteOptions", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayRecentFiles
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayRecentFiles", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayRecentFiles", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayScrollBars
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayScrollBars", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayScrollBars", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool DisplayStatusBar
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DisplayStatusBar", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("DisplayStatusBar", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool EditDirectlyInCell
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EditDirectlyInCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("EditDirectlyInCell", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool EnableAnimations
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EnableAnimations", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("EnableAnimations", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool EnableAutoComplete
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EnableAutoComplete", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("EnableAutoComplete", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlEnableCancelKey EnableCancelKey
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EnableCancelKey", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlEnableCancelKey)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("EnableCancelKey", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool EnableEvents
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EnableEvents", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("EnableEvents", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool EnableSound
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("EnableSound", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("EnableSound", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ExtendList
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ExtendList", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ExtendList", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public MsoFeatureInstall FeatureInstall
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FeatureInstall", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (MsoFeatureInstall)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FeatureInstall", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool FixedDecimal
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FixedDecimal", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FixedDecimal", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int FixedDecimalPlaces
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FixedDecimalPlaces", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("FixedDecimalPlaces", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool GenerateGetPivotData
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("GenerateGetPivotData", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("GenerateGetPivotData", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public XlMousePointer Cursor
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Cursor", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlMousePointer)returnValue; 
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Cursor", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ScreenUpdating
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ScreenUpdating", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue; 
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ScreenUpdating", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool RollZoom
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RollZoom", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("RollZoom", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool UseSystemSeparators
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UseSystemSeparators", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("UseSystemSeparators", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Visible
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Visible", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {                
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Visible", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public int Hinstance
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Hinstance", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public int Hwnd
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Hwnd", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        public double Height
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Height", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Height", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool IgnoreRemoteRequests 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("IgnoreRemoteRequests", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("IgnoreRemoteRequests", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Interactive
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Interactive", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Interactive", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Iteration
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Iteration", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Iteration", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double Left
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Left", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Left", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string LibraryPath
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LibraryPath", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("LibraryPath", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool MapPaperSize
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MapPaperSize", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("MapPaperSize", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool MathCoprocessorAvailable
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MathCoprocessorAvailable", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("MathCoprocessorAvailable", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double MaxChange
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MaxChange", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("MaxChange", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double MaxIterations
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MaxIterations", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("MaxIterations", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double MemoryFree
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MemoryFree", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
        }

        public double MemoryTotal
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MemoryTotal", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
        }

        public double MemoryUsed
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MemoryUsed", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
        }

        public bool MouseAvailable
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MouseAvailable", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool MoveAfterReturn
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MoveAfterReturn", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public XlDirection MoveAfterReturnDirection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("MoveAfterReturnDirection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlDirection)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("MoveAfterReturnDirection", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Name
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Name", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string NetworkTemplatesPath
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("NetworkTemplatesPath", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string OnWindow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("OnWindow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("OnWindow", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string OrganizationName
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("OrganizationName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string OperatingSystem
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("OperatingSystem", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string Path
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Path", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string PathSeparator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PathSeparator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public bool PivotTableSelection
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PivotTableSelection", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("PivotTableSelection", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string ProductCode
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ProductCode", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public bool PromptForSummaryInfo
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("PromptForSummaryInfo", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("PromptForSummaryInfo", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool Ready
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Ready", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public bool RecordRelative
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RecordRelative", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public XlReferenceStyle ReferenceStyle 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ReferenceStyle", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlReferenceStyle)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ReferenceStyle", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string ThousandsSeparator
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ThousandsSeparator", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ThousandsSeparator", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double SheetsInNewWorkbook
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("SheetsInNewWorkbook", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("SheetsInNewWorkbook", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowChartTipNames
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowChartTipNames", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ShowChartTipNames", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowChartTipValues
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowChartTipValues", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ShowChartTipValues", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowStartupDialog
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowStartupDialog", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ShowStartupDialog", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowToolTips
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowToolTips", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ShowToolTips", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool ShowWindowsInTaskbar
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ShowWindowsInTaskbar", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("ShowWindowsInTaskbar", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string StandardFont
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("StandardFont", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("StandardFont", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double StandardFontSize
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("StandardFontSize", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("StandardFontSize", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string StartupPath
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("StartupPath", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string TemplatesPath
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TemplatesPath", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public double Top
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Top", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Top", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string TransitionMenuKey
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TransitionMenuKey", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("TransitionMenuKey", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double TransitionMenuKeyAction
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TransitionMenuKeyAction", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("TransitionMenuKeyAction", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool TransitionNavigKeys
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("TransitionNavigKeys", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("TransitionNavigKeys", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public double UsableHeight
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UsableHeight", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
        }

        public double UsableWidth
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UsableWidth", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
        }

        public bool UserControl
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UserControl", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("UserControl", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string UserLibraryPath
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UserLibraryPath", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public string UserName
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("UserName", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("UserName", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public string Value
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Value", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (string)returnValue;
            }
        }

        public double Width
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Width", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (double)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("Width", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        public bool WindowsForPens
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WindowsForPens", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (bool)returnValue;
            }
        }

        public XlWindowState WindowState
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("WindowState", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (XlWindowState)returnValue;
            }
            set
            {
                object[] parameter = new object[1];
                parameter[0] = value;
                InstanceType.InvokeMember("WindowState", BindingFlags.SetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            }
        }

        #endregion
  
        #region ComReference Properties
        
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

        
        public XlAutoRecover AutoRecover
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AutoRecover", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlAutoRecover newClass = new XlAutoRecover(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        public XlAddins AddIns
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("AddIns", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlAddins newClass = new XlAddins(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        public XlCOMAddins COMAddIns
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("COMAddIns", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlCOMAddins newClass = new XlCOMAddins(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }

        }

        
        public XlCommandBars CommandBars
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("CommandBars", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlCommandBars newClass = new XlCommandBars(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }

        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("Dialogs"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlDialogs Dialogs
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Dialogs", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlDialogs newClass = new XlDialogs(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("Windows"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlWindows Windows
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Windows", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlWindows newClass = new XlWindows(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }


        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("Workbooks Collection in Microsoft Excel"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlWorkbooks Workbooks
        {

            get
            {
                object returnValue  = InstanceType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlWorkbooks newClass = new XlWorkbooks(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }

        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ActiveWorkbook, can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlWorkbook ActiveWorkbook
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveWorkbook", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlWorkbook newClass = new XlWorkbook(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
       
        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ThisWorkbook, can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlWorkbook ThisWorkbook
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ThisWorkbook", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlWorkbook newClass = new XlWorkbook(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ActiveWorksheet, can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
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
      
        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ActiveCell , can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlRange ActiveCell 
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("Cells , can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlRange Cells
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Cells", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        
        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ThisCell , can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlRange ThisCell
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ThisCell", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("Columns , can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlRange Columns
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Columns", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        
        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("Rows , can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlRange Rows
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Rows", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlRange newClass = new XlRange(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ActiveChart , can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlChart ActiveChart
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveChart", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlChart newClass = new XlChart(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ActiveWindow , can be NULL"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlWindow ActiveWindow
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ActiveWindow", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlWindow newClass = new XlWindow(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
 
        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("LanguageSettings"), EditorBrowsable(EditorBrowsableState.Never)]         
        public XlLanguageSettings LanguageSettings
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("LanguageSettings", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlLanguageSettings newClass = new XlLanguageSettings(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }

        }
        
        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("DefaultWebOptions"), EditorBrowsable(EditorBrowsableState.Always)]
        public XlDefaultWebOptions DefaultWebOptions
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("DefaultWebOptions", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlDefaultWebOptions newClass = new XlDefaultWebOptions(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ErrorCheckingOptions"), EditorBrowsable(EditorBrowsableState.Always)]
        public XlErrorCheckingOptions ErrorCheckingOptions
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ErrorCheckingOptions", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlErrorCheckingOptions newClass = new XlErrorCheckingOptions(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("FindFormat"), EditorBrowsable(EditorBrowsableState.Always)]
        public XlCellFormat FindFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("FindFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlCellFormat newClass = new XlCellFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }
        

        [CategoryAttribute("Application Properties"), DescriptionAttribute("ReplaceFormat"), EditorBrowsable(EditorBrowsableState.Always)]
        public XlCellFormat ReplaceFormat
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ReplaceFormat", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlCellFormat newClass = new XlCellFormat(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("ODBCErrors"), EditorBrowsable(EditorBrowsableState.Always)]
        public XlODBCErrors ODBCErrors
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("ODBCErrors", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlODBCErrors newClass = new XlODBCErrors(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("OLEDBErrors"), EditorBrowsable(EditorBrowsableState.Always)]
        public XlOLEDBErrors OLEDBErrors
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("OLEDBErrors", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlOLEDBErrors newClass = new XlOLEDBErrors(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("Parent"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlApplication Parent
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Parent", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlApplication newClass = new XlApplication(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        
        [CategoryAttribute("Application Properties"), DescriptionAttribute("RecentFiles"), EditorBrowsable(EditorBrowsableState.Never)]
        public XlRecentFiles RecentFiles
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("RecentFiles", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null; 
                XlRecentFiles newClass = new XlRecentFiles(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region ComReference Properties with Parameter

        public XlFileDialog FileDialog(MsoFileDialogType fileDialogType)
        {
            object[] parameter = new object[1];
            parameter[0] = fileDialogType;
            object returnValue  = InstanceType.InvokeMember("FileDialog", BindingFlags.GetProperty, null, ComReference, parameter, XlLateBindingApiSettings.XlThreadCulture);
            if (null == returnValue) return null; 
            XlFileDialog newClass = new XlFileDialog(this, returnValue);
            ListChildReferences.Add(newClass);
            return newClass;
        }
        
        #endregion

        #region Methods
        
        public int DDEInitiate(string app, string topic)
        {
            object[] paramArray = new object[2];
            paramArray[0] = app;
            paramArray[1] = topic;
            int returnValue = (int)InstanceType.InvokeMember("DDEInitiate", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return returnValue; 
        }

        public void DDERequest(int channel, string String)
        {
            object[] paramArray = new object[2];
            paramArray[0] = channel;
            paramArray[1] = String;
            InstanceType.InvokeMember("DDERequest", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void DDETerminate(int channel)
        {
            object[] paramArray = new object[1];
            paramArray[0] = channel;
            InstanceType.InvokeMember("DDETerminate", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void DDEExecute(int channel, string String)
        {
            object[] paramArray = new object[2];
            paramArray[0] = channel;
            paramArray[1] = String;
            InstanceType.InvokeMember("DDEExecute", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void ActivateMicrosoftApp(XlMSApplication index)
        {
            object[] paramArray = new object[1];
            paramArray[0] = index;
            InstanceType.InvokeMember("ActivateMicrosoftApp", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void DeleteChartAutoFormat(string name)
        {
            object[] paramArray = new object[1];
            paramArray[0] = name;
            InstanceType.InvokeMember("DeleteChartAutoFormat", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void DeleteCustomList(int listNum)
        {
            object[] paramArray = new object[1];
            paramArray[0] = listNum;
            InstanceType.InvokeMember("DeleteCustomList", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Evaluate(object name)
        {
            object[] paramArray = new object[1];
            paramArray[0] = name;
            InstanceType.InvokeMember("Evaluate", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void DoubleClick()
        {
            InstanceType.InvokeMember("DoubleClick", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Calculate()
        {
            InstanceType.InvokeMember("Calculate", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }
 
        public double CentimetersToPoints(Double centimeters)
        {
            object[] paramArray = new object[1];
            paramArray[0] = centimeters;
            object returnValue  = InstanceType.InvokeMember("CentimetersToPoints", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return(double)returnValue; 
        }

        public double InchesToPoints(double inches)
        {
            object[] paramArray = new object[1];
            paramArray[0] = inches;
            object returnValue  = InstanceType.InvokeMember("InchesToPoints", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
            return(double)returnValue; 
        }

        public void Quit()
        {            
            InstanceType.InvokeMember("Quit", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);              
        }

        #endregion

        #region Events

        public event AppEvents_NewWorkbookEventHandler NewWorkbook;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseNewWorkbookEvent(object __p1)
        {
            if (null == NewWorkbook)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);            
        }

        public event AppEvents_SheetActivateEventHandler SheetActivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetActivateEvent(object __p1)
        {
            if (null == SheetActivate)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);
            SheetActivate(ws);
        }

        public event AppEvents_SheetBeforeDoubleClickEventHandler SheetBeforeDoubleClick;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetBeforeDoubleClickEvent(object __p1, object __p2, ref bool Cancel)
        {
            if (null == SheetBeforeDoubleClick)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);

            XlRange ra = new XlRange(this, __p2);
            ListChildReferences.Add(ra);

            SheetBeforeDoubleClick(ws, ra, ref Cancel);
        }

        public event AppEvents_SheetBeforeRightClickEventHandler SheetBeforeRightClick;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetBeforeRightClickEvent(object __p1, object __p2, ref bool Cancel)
        {
            if (null == SheetBeforeRightClick)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);

            XlRange ra = new XlRange(this, __p2);
            ListChildReferences.Add(ra);

            SheetBeforeRightClick(ws, ra, ref Cancel);
        }

        public event AppEvents_SheetCalculateEventHandler SheetCalculate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetCalculateEvent(object __p1)
        {
            if (null == SheetCalculate)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);
            SheetCalculate(ws);
        }
    
        public event AppEvents_SheetChangeEventHandler SheetChange;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetChangeEvent(object __p1, object __p2)
        {
            if (null == SheetChange)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);

            XlRange ra = new XlRange(this, __p2);
            ListChildReferences.Add(ra);

            SheetChange(ws, ra);
        }

        public event AppEvents_SheetDeactivateEventHandler SheetDeactivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetDeactivateEvent(object __p1)
        {
            if (null == SheetDeactivate)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);
            SheetDeactivate(ws);
        }

        public event AppEvents_SheetFollowHyperlinkEventHandler SheetFollowHyperlink;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetFollowHyperlinkEvent(object __p1, object __p2)
        {
            if (null == SheetFollowHyperlink)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);

            XlHyperlink hl = new XlHyperlink(this, __p2);
            ListChildReferences.Add(hl);

            SheetFollowHyperlink(ws, hl);
        }

        public event AppEvents_SheetPivotTableUpdateEventHandler SheetPivotTableUpdate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetPivotTableUpdateEvent(object __p1, object __p2)
        {
            if (null == SheetPivotTableUpdate)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);

            XlPivotTable pt = new XlPivotTable(this, __p2);
            ListChildReferences.Add(pt);

            SheetPivotTableUpdate(ws, pt);
        }

        public event AppEvents_SheetSelectionChangeEventHandler SheetSelectionChange;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseSheetSelectionChangeEvent(object __p1, object __p2)
        {
            if (null == SheetSelectionChange)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorksheet ws = new XlWorksheet(this, __p1);
            ListChildReferences.Add(ws);

            XlRange ra = new XlRange(this, __p2);
            ListChildReferences.Add(ra);

            SheetSelectionChange(ws, ra);
        }

        public event AppEvents_WindowActivateEventHandler WindowActivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWindowActivateEvent(object __p1, object __p2)
        {
            if (null == WindowActivate)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            XlWindow wi = new XlWindow(this, __p2);
            ListChildReferences.Add(wi);

            WindowActivate(wb, wi);
        }

        public event AppEvents_WindowDeactivateEventHandler WindowDeactivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWindowDeactivateEvent(object __p1, object __p2)
        {
            if (null == WindowDeactivate)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            XlWindow wi = new XlWindow(this, __p2);
            ListChildReferences.Add(wi);

            WindowDeactivate(wb, wi);
        }

        public event AppEvents_WindowResizeEventHandler WindowResize;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWindowResizeEvent(object __p1, object __p2)
        {
            if (null == WindowResize)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            XlWindow wi = new XlWindow(this, __p2);
            ListChildReferences.Add(wi);

            WindowResize(wb, wi);
        }

        public event AppEvents_WorkbookActivateEventHandler WorkbookActivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookActivateEvent(object __p1)
        {
            if (null == WorkbookActivate)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookActivate(wb);
        }

        public event AppEvents_WorkbookAddinInstallEventHandler WorkbookAddinInstall;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookAddinInstallEvent(object __p1)
        {
            if (null == WorkbookAddinInstall)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookAddinInstall(wb);
        }

        public event AppEvents_WorkbookAddinUninstallEventHandler WorkbookAddinUninstall;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookAddinUninstallEvent(object __p1)
        {
            if (null == WorkbookAddinUninstall)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookAddinUninstall(wb);
        }

        public event AppEvents_WorkbookBeforeCloseEventHandler WorkbookBeforeClose;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookBeforeCloseEvent(object __p1, ref bool Cancel)
        {
            if (null == WorkbookBeforeClose)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookBeforeClose(wb, ref Cancel);
        }

        public event AppEvents_WorkbookBeforePrintEventHandler WorkbookBeforePrint;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookBeforePrintEvent(object __p1, ref bool Cancel)
        {
            if (null == WorkbookBeforePrint)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookBeforePrint(wb, ref Cancel);
        }

        public event AppEvents_WorkbookBeforeSaveEventHandler WorkbookBeforeSave;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookBeforeSaveEvent(object __p1, bool SaveAsUI, ref bool Cancel)
        {
            if (null == WorkbookBeforeSave)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookBeforeSave(wb, SaveAsUI, ref Cancel);
        }

        public event AppEvents_WorkbookDeactivateEventHandler WorkbookDeactivate;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookDeactivateEvent(object __p1)
        {
            if (null == WorkbookDeactivate)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookDeactivate(wb);
        }

        public event AppEvents_WorkbookNewSheetEventHandler WorkbookNewSheet;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookNewSheetEvent(object __p1, object __p2)
        {
            if (null == WorkbookNewSheet)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            XlWorksheet ws = new XlWorksheet(this, __p2);
            ListChildReferences.Add(ws);

            WorkbookNewSheet(wb, ws);
        }

        public event AppEvents_WorkbookOpenEventHandler WorkbookOpen;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookOpenEvent(object __p1)
        {
            if (null == WorkbookOpen)
            {
                Marshal.ReleaseComObject(__p1);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            WorkbookOpen(wb);
        }

        public event AppEvents_WorkbookPivotTableCloseConnectionEventHandler WorkbookPivotTableCloseConnection;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookPivotTableCloseConnectionEvent(object __p1, object __p2)
        {
            if (null == WorkbookPivotTableCloseConnection)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            XlPivotTable pt = new XlPivotTable(this, __p2);
            ListChildReferences.Add(pt);

            WorkbookPivotTableCloseConnection(wb, pt);
        }

        public event AppEvents_WorkbookPivotTableOpenConnectionEventHandler WorkbookPivotTableOpenConnection;
        [EnvironmentPermissionAttribute(SecurityAction.LinkDemand, Unrestricted = true)]
        internal void RaiseWorkbookPivotTableOpenConnectionEvent(object __p1, object __p2)
        {
            if (null == WorkbookPivotTableOpenConnection)
            {
                Marshal.ReleaseComObject(__p1);
                Marshal.ReleaseComObject(__p2);
                return;
            }

            XlWorkbook wb = new XlWorkbook(this, __p1);
            ListChildReferences.Add(wb);

            XlPivotTable pt = new XlPivotTable(this, __p2);
            ListChildReferences.Add(pt);

            WorkbookPivotTableOpenConnection(wb, pt);
        }

        #endregion

        #region IXlEventBinding Members

        public void SetupEventBinding()
        {
            _eventBridge = new XlApplicationEvents();
            _eventBridge.SetupEventBinding(this);
        }

        public void RemoveEventBinding()
        {
            if(null!=_eventBridge) 
                _eventBridge.RemoveEventBinding();
        }

        #endregion
    }
}
