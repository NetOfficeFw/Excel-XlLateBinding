using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;


namespace LateBindingApi.Excel.Shapes
{
    public class XlDiagramNodes : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlDiagramNodes(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
		}

        #endregion
 
        #region Methods

        public void SelectAll()
        {
            InstanceType.InvokeMember("SelectAll", BindingFlags.InvokeMethod, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
        }

        #endregion
 
        #region Scalar Properties

        public int Count
        {
            get
            {
                object returnValue  = InstanceType.InvokeMember("Count", BindingFlags.GetProperty, null, ComReference, null, XlLateBindingApiSettings.XlThreadCulture);
                return (int)returnValue;
            }
        }

        #endregion
         
        #region ComReference Properties

        /// <summary>
        /// returns a XlDiagramNode by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlDiagramNode this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlDiagramNode newClass = new XlDiagramNode(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region ForEach

        /// <summary>
        /// Foreach Enumerator
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
  
            int iCount = Count;
            XlDiagramNode[] res_shapes = new XlDiagramNode[iCount];

            for (int i = 1; i <= iCount; i++)
                res_shapes[i - 1] = this[i];

            for (int i = 0; i < res_shapes.Length; i++)
            {
                yield return res_shapes[i];
            }


        }

        #endregion
    }
}
