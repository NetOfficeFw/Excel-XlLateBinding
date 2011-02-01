using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Globalization;
using System.IO;

using XlLateBinding.Debug;
using XlLateBinding.Enum;
using XlLateBinding.Exceptions;

namespace XlLateBinding
{
    /// <summary>
    /// Represents the Excel Application, not bind to a specific Version
    /// </summary>
    public class Application : IDisposable
    {
       
        #region Constants
        
        // Default Thread Culture to Excel Application
        private const string _DefaultCulture = "en-US";

        // Default Progid from Excel without Version
        private const string _DefaultProgId = "Excel.Application";

        #endregion

        #region Member

        private bool                _IsDisposed           = false;

        private object              _Application          = null;

        private CultureInfo         _cultureInfo          = null;
        private Type                _Type                 = null;       
        
        private LanguageSettings    _languageSettings     = null;
     
        #endregion
        
        #region Construction

        /// <summary>
        /// Create Instance means starting Excel
        /// </summary>
        /// <param name="ProgId">Specified ProgId for example: "Excel.Application.8" </param>
        public Application(string ProgId)
        {
            Initialize(ProgId); 
        }

        /// <summary>
        /// Create Instance means starting Excel
        /// </summary>
        public Application()
        {

            Initialize(_DefaultProgId); 
        }

        /// <summary>
        /// Create Instance means starting Excel
        /// </summary>
        public Application(bool bDisplayAlerts, bool bVisible)
        {

             Initialize(_DefaultProgId);

             DisplayAlerts = bDisplayAlerts;
             Visible       = bVisible;
                             
        }

        #endregion

        #region xlFramework Properties

        /// <summary>
        /// Version of xlFramework executing Assembly
        /// </summary>
        public Version FrameworkVersion
        {

            get
            {                
                return Assembly.GetExecutingAssembly().GetName().Version;               
            }

        }

        /// <summary>
        /// LCID from default used Thread by the Framework
        /// </summary>
        public string DefaultThreadLCID
        {
        
            get
            {
                return _DefaultCulture;                
            }
         
        }
 
        /// <summary>
        /// Used Thread Culture, default is DefaultThreadLCID
        /// </summary>
        public CultureInfo xlThreadCulture
        {

            get
            {                 
                return _cultureInfo;               
            }
            set
            {
                _cultureInfo = value;
            }

        }

        /// <summary>
        /// Static COM Reference to Excel
        /// </summary>
        public object xlObject
        {  

            get
            {
                 return _Application;
            }

        }

        #endregion

        #region Excel Application Properties

        /// <summary>
        /// Original LanguageSettings 
        /// </summary>
        public LanguageSettings LanguageSettings
        {

            get
            {
                return _languageSettings;
            }

        }

        /// <summary>
        /// Original Property Version
        /// </summary>
        public string Version
        {
         
            get
            {                 
                  return (string)_Application.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, _Application, null, _cultureInfo);
            }
  
        }

        /// <summary>
        /// Original Property DecimalSeparator
        /// </summary>
        public string DecimalSeparator
        {

            get
            {
                    
                    return (string)_Application.GetType().InvokeMember("DecimalSeparator", BindingFlags.GetProperty, null, _Application, null, _cultureInfo);
                    
            }
            set
            {

                    object[] parameter = new object[1];
                    parameter[0] = value;
                    _Application.GetType().InvokeMember("DecimalSeparator", BindingFlags.SetProperty, null, _Application, parameter, _cultureInfo);
             
            }        

        }

        /// <summary>
        /// Original Property DisplayAlerts 
        /// </summary>
        public bool DisplayAlerts 
        {

            get 
            {
                return (bool)_Application.GetType().InvokeMember("DisplayAlerts", BindingFlags.GetProperty, null, _Application, null, _cultureInfo);
            }
            set
            {
          
                object[] parameter = new object[1];
                parameter[0] = value;
                _Application.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, _Application, parameter, _cultureInfo);                
            }        
            
        }

        /// <summary>
        /// Original Property Cursor 
        /// </summary>
        public XLMousePointer Cursor
        {

            get
            {
                    return (XLMousePointer)_Application.GetType().InvokeMember("Cursor", BindingFlags.GetProperty, null, _Application, null, _cultureInfo); 
            }
            set
            {
                    object[] parameter = new object[1];
                    parameter[0] = value;
                    _Application.GetType().InvokeMember("Cursor", BindingFlags.SetProperty, null, _Application, parameter, _cultureInfo);
             }

        }

        /// <summary>
        /// Original Property ScreenUpdating 
        /// </summary>
        public bool ScreenUpdating
        {

            get
            {

                return (bool)_Application.GetType().InvokeMember("ScreenUpdating", BindingFlags.GetProperty, null, _Application, null, _cultureInfo);
                
            }
            set
            {
                
                object[] parameter = new object[1];
                parameter[0] = value;
                _Application.GetType().InvokeMember("ScreenUpdating", BindingFlags.SetProperty, null, _Application, parameter, _cultureInfo);
            
            }

        }

        /// <summary>
        /// Original Property UseSystemSeparators 
        /// </summary>
        public bool UseSystemSeparators
        {

            get
            {
                return (bool)_Application.GetType().InvokeMember("UseSystemSeparators", BindingFlags.GetProperty, null, _Application, null, _cultureInfo);          
            }
            set
            {                
                    object[] parameter = new object[1];
                    parameter[0] = value;
                    _Application.GetType().InvokeMember("UseSystemSeparators", BindingFlags.SetProperty, null, _Application, parameter, _cultureInfo);
            }

        }

        /// <summary>
        /// Original Property Visible from Main Window
        /// </summary>
        public bool Visible
        {
            
             get
             {
                    return (bool)_Application.GetType().InvokeMember("Visible", BindingFlags.GetProperty, null, _Application, null, _cultureInfo);               
             }
             set 
             {
                    object[] parameter = new object[1];
                    parameter[0] = value;
                    _Application.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, _Application, parameter, _cultureInfo);
              }

        }

        /// <summary>
        /// Original Property Hwnd from Main Window
        /// </summary>
        public int Hwnd
        {

            get 
            {
                return (int)_Application.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, _Application, null, _cultureInfo);               
            }
  
        }
      
        #endregion

        #region Excel Application Methods
        
        /// <summary>
        /// Original Quit Mehtod
        /// Dont forget to perform the dispose methoder after calling Quit
        /// </summary> 
        public void Quit()
        {
            _Application.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, _Application, null, _cultureInfo);              
        }

        #endregion

        #region Destruction

        /// <summary>
        /// Dispose the Excel Instance. The Excel Process dies immediately.
        /// </summary>
        /// <Exceptions>DisposeException </Exceptions>
        public void Dispose()
        {

            try
            {

                if (true == _IsDisposed)
                    throw (new DisposeException("The instance is already disposed."));
  
                if (null != _Application) 
                     Marshal.FinalReleaseComObject(_Application);                    

            }
            catch (Exception ex)
            {
                throw (ex);
            }
            finally
            {

                _cultureInfo = null;
                _Application = null;
                _Type        = null;
                _IsDisposed  = true;

            }
        }

        #endregion

        #region Initialize 

        /// <summary>
        /// Construction Helper
        /// </summary>
        private void Initialize(string ProgId)
        {

            #region GetCulture

            _cultureInfo = CultureInfo.GetCultureInfo(_DefaultCulture);
            if (null == _cultureInfo)
                throw (new InitializeException("GetCultureInfo " + _DefaultCulture + " failed."));

            #endregion
            
            #region get type
            
            _Type = Type.GetTypeFromProgID(ProgId);
            if (null == _Type)
                throw (new InitializeException("GetTypeFromProgID " + ProgId + " failed."));
           
            #endregion

            #region create application
            
            _Application = Activator.CreateInstance(_Type);
            if (null == _Application)
                throw (new InitializeException("CreateInstance " + ProgId + " failed."));
             
            #endregion

            _languageSettings = new LanguageSettings(this);   
        
        }

        #endregion

    }

}
