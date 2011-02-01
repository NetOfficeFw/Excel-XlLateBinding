using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.ComponentModel;

using LateBindingApi.Excel.Interfaces;
using LateBindingApi.Excel.Enums;

namespace LateBindingApi.Excel.Shapes
{
    public class XlShapeNodes : XlNonCreatable, System.Collections.IEnumerable
    {
        #region Construction

        internal XlShapeNodes(IXlObject parentReference, object comReference): base(parentReference, comReference)
		{
		 
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
        /// returns a Shape by Index, not 0 based
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public XlShapeNode this[int index]
        {
            get
            {
                object[] paramArray = new object[1];
                paramArray[0] = index;
                object returnValue  = InstanceType.InvokeMember("Item", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
                if (null == returnValue) return null;
                XlShapeNode newClass = new XlShapeNode(this, returnValue);
                ListChildReferences.Add(newClass);
                return newClass;
            }
        }

        #endregion

        #region Methods

        public void Insert(int index, MsoSegmentType segmentType, MsoEditingType editingType, Single x1, Single y1)
        {
            object[] paramArray = new object[9];
            paramArray[0] = index;
            paramArray[1] = segmentType;
            paramArray[2] = editingType;
            paramArray[3] = x1;
            paramArray[4] = y1;
            paramArray[5] = Missing.Value;
            paramArray[6] = Missing.Value;
            paramArray[7] = Missing.Value;
            paramArray[8] = Missing.Value;
            InstanceType.InvokeMember("Insert", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
        }

        public void Insert(int index, MsoSegmentType segmentType, MsoEditingType editingType, Single x1, Single y1, Single x2, Single y2, Single x3, Single y3)
        {
            object[] paramArray = new object[9];
            paramArray[0] = index;
            paramArray[1] = segmentType;
            paramArray[2] = editingType;
            paramArray[3] = x1;
            paramArray[4] = y1;
            paramArray[5] = x2;
            paramArray[6] = y2;
            paramArray[7] = x3;
            paramArray[8] = y3;
            InstanceType.InvokeMember("Insert", BindingFlags.InvokeMethod, null, ComReference, paramArray, XlLateBindingApiSettings.XlThreadCulture);
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
            XlShapeNode[] res_shapes = new XlShapeNode[iCount];

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
