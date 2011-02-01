using System;
using System.Windows.Forms; 
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Extensibility;

using Excel = LateBindingApi.Excel;
using Office = LateBindingApi.Office;

using LateBindingApi.Excel.Enums;
using LateBindingApi.Office.Enums; 

namespace COMAddinExample
{
    #region Ribbon Interfaces
    [ComImport, Guid("000C03A7-0000-0000-C000-000000000046"), TypeLibType((short)0x1040)]
    public interface IRibbonUI
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        void Invalidate();
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
        void InvalidateControl([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
        void InvalidateControlMso([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
        void ActivateTab([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(5)]
        void ActivateTabMso([In, MarshalAs(UnmanagedType.BStr)] string ControlID);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(6)]
        void ActivateTabQ([In, MarshalAs(UnmanagedType.BStr)] string ControlID, [In, MarshalAs(UnmanagedType.BStr)] string Namespace);
    }

    [ComImport, Guid("000C0395-0000-0000-C000-000000000046"), TypeLibType((short)0x1040)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)] get; }
        [DispId(2)]
        object Context { [return: MarshalAs(UnmanagedType.IDispatch)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)] get; }
        [DispId(3)]
        string Tag { [return: MarshalAs(UnmanagedType.BStr)] [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)] get; }
    }

    [ComImport, TypeLibType((short)0x1040), Guid("000C0396-0000-0000-C000-000000000046")]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(1)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    #endregion

    [ComVisible(true)]
    [GuidAttribute("05B31000-90EF-4a62-A67C-77C2526D7364"), ProgId("COMAddinExample.COMAddin")]
    public class XlCOMAddin : IDTExtensibility2, IRibbonExtensibility
    {
        private static readonly string _prodId = "COMAddinExample.COMAddin";
       
        Excel.Application _excelApplication;

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
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string GetSubKeyName(Type type)
        {
            return @"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable";
        }

        private static void DeleteExcelAddInKey()
        {
            Registry.CurrentUser.DeleteSubKey("Software\\Microsoft\\Office\\Excel\\AddIns\\" + _prodId);
        }

        private static void CreateExcelAddInKey()
        {
            RegistryKey rk;
            Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\Office\\Excel\\AddIns\\" + _prodId);
            rk = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office\\Excel\\AddIns\\" + _prodId, true);
            rk.SetValue("LoadBehavior", Convert.ToInt32(3));
            rk.SetValue("FriendlyName", "ComAddinExample");
            rk.SetValue("Description", "LateBindingApi ComAddinExample");
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
            try
            {
                _excelApplication = new Excel.Application(null, Application);
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != _excelApplication)
                    _excelApplication.Dispose();
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            
        }

        #endregion

        #region IRibbonExtensibility Members

        public string GetCustomUI(string RibbonID)
        {           
            return ReadTextFileFromRessource("RibbonUI.xml");
        }
  
        #endregion

        #region Gui Trigger

        public void OnAction(IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    case "customButton1":
                        MessageBox.Show("This is the first sample button.");
                        break;
                    case "customButton2":
                        MessageBox.Show("This is the second sample button.");
                        break;
                    default:
                        MessageBox.Show("Unkown Control Id: " + control.Id);
                        break;
                }
            }
            catch (Exception throwedException)
            {
                string details = string.Format("{1}{1}Details:{1}{1}{0}", throwedException.Message, Environment.NewLine);
                MessageBox.Show("An error occured." + details, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        #endregion

        #region Private Helper

        private static string ReadTextFileFromRessource(string fileName)
        {
            fileName = "COMAddinExample." + fileName;

            System.IO.Stream ressourceStream;
            System.IO.StreamReader textStreamReader;
            try
            {
                ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                ressourceStream.Close();
                textStreamReader.Close();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }
        }

        #endregion
    }
}
