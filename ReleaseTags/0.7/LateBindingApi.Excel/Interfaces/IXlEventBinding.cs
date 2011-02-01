using System;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Excel.Interfaces
{
    public interface IXlEventBinding
    {
        void SetupEventBinding();
        void RemoveEventBinding();   
    }

}
