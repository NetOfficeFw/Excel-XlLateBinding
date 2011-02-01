using System;

namespace LateBindingApi.Excel.Interfaces
{
    /// <summary>
    ///  Defines methods to manage COM References
    /// </summary>
    public interface IXlObject : IDisposable
    {
        /// <summary>
        /// The COM Proxy from COM Runtime System
        /// </summary>
        object COMReference { get; }

        /// <summary>
        /// The Instance they has been created this instance
        /// </summary>
        IXlObject ParentReference { get; set; }

        /// <summary>
        /// Release COM Reference from COM Runtime System
        /// </summary>
        void ReleaseCOMReference();

        /// <summary>
        ///  Remove the child reference from this instance collection without a release
        /// </summary>
        /// <param name="comReference"></param>
        void RemoveChildReference(IXlObject comReference);

        /// <summary>
        /// Release and remove all child references
        /// </summary>
        void ReleaseChildReferences();
    }
}
