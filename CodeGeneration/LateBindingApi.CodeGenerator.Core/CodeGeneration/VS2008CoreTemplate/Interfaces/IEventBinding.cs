﻿// Generated by LateBindingApi.CodeGenerator
using System;
using System.ComponentModel;
namespace LateBindingApi.Core
{
    public interface IEventBinding
    {
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void CallEvent(string name, object[] paramArray);

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        void DisposeSinkHelper();
    }
}