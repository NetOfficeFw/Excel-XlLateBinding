using System;

namespace LateBindingApi.Excel.Interfaces
{
    /// <summary>
    ///  Defines methods to manage COM References for a creatable class
    /// </summary>
    public interface IXlCreatable : IXlObject 
    {
        /// <summary>
        /// Creates a COM Reference for the class. 
        /// </summary>
        /// <returns></returns>
        void CreateCOMReference(string progId);
    }
}
