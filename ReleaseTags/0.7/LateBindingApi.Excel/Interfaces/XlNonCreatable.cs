using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace LateBindingApi.Excel.Interfaces
{
    public abstract class XlNonCreatable : IXlNonCreatable
    {
        #region Fields

        private Type _InstanceType;
        private object _ComReference;
        private IXlObject _ParentReference;
        private List<IXlObject> _ListChildReferences = new List<IXlObject>();

        #endregion

        #region Properties

        internal Type InstanceType
        {
            get 
            {
                return _InstanceType;
            }
        }

        internal object ComReference
        {
            get
            {
                return _ComReference;
           
            }
        }

        internal List<IXlObject> ListChildReferences
        {
            get
            {
                return _ListChildReferences;

            }
        }

        #endregion

        #region Construction

        /// <summary>
        /// IXlNonCreatable Constructor
        /// </summary>
        /// <param name="parentReference"></param>
        /// <param name="comReference"></param>
        internal XlNonCreatable(IXlObject parentReference, object comReference)
		{
			_ParentReference = parentReference;
			_ComReference = comReference;
            _InstanceType = _ComReference.GetType();

            // in case of this is a type with event support we enable the binding to COM event point
            IXlEventBinding typeEvent = this as IXlEventBinding;
            if (typeEvent != null)
            {
                typeEvent.SetupEventBinding();
            }
		}        

        #endregion

        #region IXlObject

        
        public object COMReference
        {
            get
            {
                return _ComReference;
            }
        }

        
        public IXlObject ParentReference
        {
            get
            {
                return _ParentReference;
            }
            set
            {
                _ParentReference = value;
            }
        }

        public void ReleaseCOMReference()
        {
            this.RemoveEventBinding();

            // remove himself from parent childlist
            if (null != _ParentReference)
            {
                _ParentReference.RemoveChildReference(this);
                _ParentReference = null;
            }
          
           
            // finally release himself
            if (null != _ComReference)
            {
                Marshal.ReleaseComObject(_ComReference);
                _ComReference = null;
            }
        }

        public void RemoveChildReference(IXlObject comReference)
        {
            ListChildReferences.Remove(comReference);
        }

        public void ReleaseChildReferences()
        {
            // release all childs
            foreach (IXlObject itemReference in _ListChildReferences)
            {
                itemReference.ParentReference = null;
                itemReference.ReleaseChildReferences();
                itemReference.ReleaseCOMReference();
            }

            // and now clear
            ListChildReferences.Clear();
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            this.RemoveEventBinding();
            this.ReleaseChildReferences();
            this.ReleaseCOMReference();
            GC.SuppressFinalize(this);
        }

        #endregion

        #region EventBinding

        private void RemoveEventBinding()
        {
            // in case of this is a type with event support we remove the binding to COM event point
            IXlEventBinding typeEvent = this as IXlEventBinding;
            if (typeEvent != null)
            {
                typeEvent.RemoveEventBinding();
            }
        }
        
        #endregion
    }
}
