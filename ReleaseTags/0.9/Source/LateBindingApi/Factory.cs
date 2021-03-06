﻿using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using COMTypes = System.Runtime.InteropServices.ComTypes;

namespace LateBindingApi.Core
{
    [Guid("00020400-0000-0000-c000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IDispatch
    {
        int GetTypeInfoCount();
        System.Runtime.InteropServices.ComTypes.ITypeInfo GetTypeInfo([MarshalAs(UnmanagedType.U4)] int iTInfo,[MarshalAs(UnmanagedType.U4)] int lcid);

        [PreserveSig]
        int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames, int cNames, int lcid, [MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        
        [PreserveSig]
        int Invoke(int dispIdMember, ref Guid riid, [MarshalAs(UnmanagedType.U4)] int lcid, [MarshalAs(UnmanagedType.U4)]  
                                                        int dwFlags, ref System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, 
                                                        [Out, MarshalAs(UnmanagedType.LPArray)] object[] pVarResult, 
                                                        ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, 
                                                        [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
    }

    public static class Factory
    {
        #region Fields

        private static List<IFactoryInfo> _factoryList = new List<IFactoryInfo>();

        private static Dictionary<string, Type> _typeCache = new Dictionary<string, Type>();

        #endregion
        
        #region COMObject
  
        /// <summary>
        /// creates a new COMObject based on classType of comProxy 
        /// </summary>
        /// <param name="caller"></param>
        /// <param name="comProxy"></param>
        /// <returns></returns>
        public static COMObject CreateObjectFromComProxy(COMObject caller, object comProxy)
        { 
            if (null == comProxy)
                return null;
           
            IFactoryInfo factoryInfo = GetFactoryInfo(comProxy);
           
            string className = TypeDescriptor.GetClassName(comProxy);
            string fullClassName = factoryInfo.Namespace + "." + factoryInfo.Prefix + className;

            // create new classType
            Type comProxyType = comProxy.GetType();
            COMObject newObject = CreateObjectFromComProxy(factoryInfo, caller, comProxy, comProxyType, className, fullClassName);
            return newObject;
        }

        private static COMObject CreateObjectFromComProxy(IFactoryInfo factoryInfo, COMObject caller, object comProxy, Type comProxyType, string className, string fullClassName)
        {
            Type classType = null;
            if (true == _typeCache.TryGetValue(fullClassName, out classType))
            {
                // cached classType
                object newClass = Activator.CreateInstance(classType, new object[] { caller, comProxy });
                return (COMObject)newClass;
            }
            else
            {
                // create new classType
                classType = factoryInfo.Assembly.GetType(fullClassName);
                if (null == classType)
                    throw new ArgumentException("Class not exists: " + fullClassName);

                _typeCache.Add(fullClassName, classType);
                object newClass = Activator.CreateInstance(classType, new object[] { caller, comProxy, comProxyType });
                return (COMObject)newClass;
            }
        }

        #endregion

        #region COMVariant

        /// <summary>
        ///  creates a new COMVariant based on type of comVariant 
        /// </summary>
        /// <param name="caller"></param>
        /// <param name="comVariant"></param>
        /// <returns></returns>
        public static COMVariant CreateVariantFromComProxy(COMObject caller, object comVariant)
        {
            if (null == comVariant)
                return null;
            
            Type comVariantType = comVariant.GetType();
            if (false == comVariantType.IsCOMObject)
            {
                COMVariant newVariant = new COMVariant(caller, comVariant, comVariantType);
                return newVariant;
            }
            else
            {
                IFactoryInfo factoryInfo = GetFactoryInfo(comVariant);
                string className = TypeDescriptor.GetClassName(comVariant);
                string fullClassName = factoryInfo.Namespace + "." + factoryInfo.Prefix + className;
                COMObject newObject = CreateObjectFromComProxy(factoryInfo, caller, comVariant, comVariantType, className, fullClassName);
                return newObject;
            }
        }

        #endregion

        /// <summary>
        /// Must be called from client assembly for COMVariant and COMObject Support
        /// Recieve FactoryInfos from all loaded LateBindingApi based Assemblies
        /// </summary>
        public static void Initialize()
        {
            _factoryList.Clear();
            Assembly callingAssembly = System.Reflection.Assembly.GetCallingAssembly();
            foreach (AssemblyName item in callingAssembly.GetReferencedAssemblies())
            {
                Assembly itemAssembly = Assembly.Load(item);
                object[] attributes = itemAssembly.GetCustomAttributes(true);
                foreach (object itemAttribute in attributes)
                {
                    string fullnameAttribute = itemAttribute.GetType().FullName;
                    if (fullnameAttribute == "LateBindingApi.Core.LateBindingAttribute")
                    {
                        Type factoryInfoType = itemAssembly.GetType(item.Name + ".Utils.FactoryInfo");
                        IFactoryInfo factoryInfo = Activator.CreateInstance(factoryInfoType) as IFactoryInfo;
                        _factoryList.Add(factoryInfo);
                    }
                }
            }
        }

        private static Guid GetParentLibGuid(object comProxy)
        {
            Guid returnGuid = Guid.Empty; 

            IDispatch dispatcher = (IDispatch)comProxy;
            COMTypes.ITypeInfo typeInfo = dispatcher.GetTypeInfo(0, 0);
            COMTypes.ITypeLib parentTypeLib = null;
          
            int i = 0;
            typeInfo.GetContainingTypeLib(out parentTypeLib, out i);
            
            IntPtr attributesPointer = IntPtr.Zero;
            parentTypeLib.GetLibAttr(out attributesPointer);

            COMTypes.TYPEATTR attributes = (COMTypes.TYPEATTR)Marshal.PtrToStructure(attributesPointer, typeof(COMTypes.TYPEATTR));
            returnGuid = attributes.guid; 

            parentTypeLib.ReleaseTLibAttr(attributesPointer);
            Marshal.ReleaseComObject(parentTypeLib);
            Marshal.ReleaseComObject(typeInfo);

            return returnGuid;
        }

        private static IFactoryInfo GetFactoryInfo(object comProxy)
        {
            Guid targetGuid = GetParentLibGuid(comProxy);
        
            foreach (IFactoryInfo item in _factoryList)
            {
                if (true == targetGuid.Equals(item.ComponentGuid))
                    return item;

            }
            
            throw new ArgumentException(TypeDescriptor.GetClassName(comProxy) + " class not found in LateBindingApi.");
        }
    }
}
