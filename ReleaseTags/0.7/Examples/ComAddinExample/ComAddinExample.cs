using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms; 
using Microsoft.Win32;
using System.Runtime.InteropServices;
using Extensibility;
using LateBindingApi.Excel;

namespace COMAddinExample    
{
    [ComVisible(true)]
    [GuidAttribute("05B31000-90EF-4a62-A67C-77C2526D7364"), ProgId("COMAddinExample.COMAddin")]
    public class XlCOMAddin : IDTExtensibility2
    {
        private static readonly string _ProdId = "COMAddinExample.COMAddin";
        XlApplication _application;

        #region COM Functions

        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.CreateSubKey(GetSubKeyName(type));
                CreateExcelAddInKey();
            }
            catch (Exception ex) 
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);     
            }

        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                Registry.ClassesRoot.DeleteSubKey(GetSubKeyName(type), false);
                DeleteExcelAddInKey();
            }
            catch (Exception ex)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", ex.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private static string GetSubKeyName(Type type)
        {
            string s = @"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable";
            return s;
        }

        private static void DeleteExcelAddInKey()
        {
            Registry.CurrentUser.DeleteSubKey("Software\\Microsoft\\Office\\Excel\\AddIns\\" + _ProdId);
        }

        private static void CreateExcelAddInKey()
        {
            RegistryKey rk;
            Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Office\\Excel\\AddIns\\" + _ProdId);
            rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office\\Excel\\AddIns\\" + _ProdId, true);
            rk.SetValue("LoadBehavior", Convert.ToInt32(3));
            rk.SetValue("FriendlyName", "ComAddinExample");
            rk.SetValue("Description", "ComAddinExample Addin");
            rk.Close();            
        }

        #endregion

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
           
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
             
        }

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _application = new XlApplication(null, Application);
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (null != _application)
                _application.Dispose();
        }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            
        }

        #endregion
    }
}
