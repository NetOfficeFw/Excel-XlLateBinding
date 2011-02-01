using System;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.CodeGenerator.Core
{  
    #region ComponentFile
    
    public class ComponentFile
    {
        #region Fields

        string _filename;
        bool _isExternal;

        #endregion

        #region Construction

        public ComponentFile()
        {
        }

        public ComponentFile(string fileName, bool isExternal)
        {
            _filename = fileName;
            _isExternal = isExternal;
        }

        #endregion

        #region Fields

        public string Filename
        {
            get
            {
                return _filename;
            }
            set
            {
                _filename = value;
            }
        }

        public bool IsExternal
        {
            get
            {
                return _isExternal;
            }
            set
            {
                _isExternal = value;
            }
        }

        #endregion
    }
    
    #endregion

    public class COMComponentReaderSettings
    {
        #region Fields

        bool _createNewProject;
        ComponentFile[] _files;

        #endregion

        #region Properties

        public ComponentFile[] Files
        {
            get
            {
                return _files;
            }
            internal set
            {
                _files = value;
            }
        }

        public bool CreateNewProject
        {
            get
            {
                return _createNewProject;
            }
            internal set
            {
                _createNewProject = value;
            }
        }

        #endregion
    }
}
