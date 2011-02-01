﻿// Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Core
{
    public static class Invoker
    {
        #region Method

        public static void Method(IObject comObject, string name, object[] paramArray)
        {
            paramArray = ValidateParamArray(paramArray);
            comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.COMProxy, paramArray, Settings.ThreadCulture);
        }

        public static object MethodReturn(IObject comObject, string name, object[] paramArray)
        {
            paramArray = ValidateParamArray(paramArray);
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.InvokeMethod, null, comObject.COMProxy, paramArray, Settings.ThreadCulture);
            return returnValue;
        }

        #endregion

        #region Property

        public static object PropertyGet(IObject comObject, string name)
        {
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.COMProxy, null, Settings.ThreadCulture);
            return returnValue;
        }

        public static object PropertyGet(IObject comObject, string name, object[] paramArray)
        {
            paramArray = ValidateParamArray(paramArray);
            object returnValue = comObject.InstanceType.InvokeMember(name, BindingFlags.GetProperty, null, comObject.COMProxy, paramArray, Settings.ThreadCulture);
            return returnValue;
        }

        public static void PropertySet(IObject comObject, string name, object value)
        {
            value = ValidateParam(value);
            comObject.InstanceType.InvokeMember(name, BindingFlags.SetProperty, null, comObject.COMProxy, new object[] { value }, Settings.ThreadCulture);
        }

        #endregion

        #region Methods

        public static object ValidateParam(object param)
        {
            if (null != param)
            {
                IObject comObject = param as IObject;
                if (null != comObject)
                    param = comObject.COMProxy;

                return param;
            }
            else
                return null;
        }

        public static object[] ValidateParamArray(object[] paramArray)
        {
            if (null != paramArray)
            {
                int parramArrayCount = paramArray.Length - 1;
                for (int i = 0; i < parramArrayCount; i++)
                {
                    paramArray[i] = ValidateParam(paramArray[i]);
                }
                return paramArray;
            }
            else
                return null;
        }

        public static void ReleaseParam(object param)
        {
            if (null != param)
            {
                Type paramType = param.GetType();
                if (true == paramType.IsCOMObject)
                    Marshal.ReleaseComObject(param);
            }
        }

        public static void ReleaseParamArray(object[] paramArray)
        {
            if (null != paramArray)
            {
                int parramArrayCount = paramArray.Length - 1;
                for (int i = 0; i < parramArrayCount; i++)
                {
                    ReleaseParam(paramArray[i]);
                }
            }
        }

        #endregion
    }
}