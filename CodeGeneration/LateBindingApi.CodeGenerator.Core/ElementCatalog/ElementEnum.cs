using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.CodeGenerator.Core
{
    public class ElementEnum
    {
        #region Fields

        private string _name;
        private List<COMEnum> _enumList = new List<COMEnum>();
        
        #endregion

        #region Properties

        public string Name
        {
            get 
            {
                return _name;
            }
        }

        public COMEnum this[int i]
        {
            get
            {
                return _enumList[i];
            }
        }

        public IEnumerator GetEnumerator()
        {
            int enumCount = Count;
            COMEnum[] returnEnums = new COMEnum[enumCount];

            for (int i = 0; i < enumCount; i++)
                returnEnums[i] = this[i];

            for (int i = 0; i < returnEnums.Length; i++)
            {
                yield return returnEnums[i];
            }

        }


        internal int Count
        {
            get 
            {
                return _enumList.Count;
            }
        }

        #endregion

        #region Construction

        internal ElementEnum(COMEnum newEnum)
        {
            _name = newEnum.Name;
            _enumList.Add(newEnum);
        }

        #endregion

        #region Methods

        internal void Add(COMEnum newEnum)
        {
            foreach (COMEnum item in _enumList)
            {
                if (item.Component == newEnum.Component)
                {
                    return;
                }
            }
            _enumList.Add(newEnum);

        }
        
        #endregion
    }
}
