using System.Xml;
using TLI;
using System.Runtime.InteropServices;

namespace LateBindingApi.CodeGenerator.Core
{
    internal partial class AliasHandler
    {
        #region Fields

        COMComponentReader _parent;

        #endregion

        #region Construction

        internal AliasHandler(COMComponentReader parent)        
        {
            _parent = parent;
        }

        #endregion

    }
}
