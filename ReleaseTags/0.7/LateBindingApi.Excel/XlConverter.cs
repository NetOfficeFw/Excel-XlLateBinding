using System;
using System.Collections.Generic;
using System.Text;

namespace LateBindingApi.Excel
{

    /// <summary>
    /// Provides various conversion methods
    /// </summary>
    public static class XlConverter
    {
        private static string[] _columnIndex = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
        
        /// <summary>
        /// Translate cell coordinates to a valid cell adress
        /// </summary>
        /// <param name="columnIndex">column index</param>
        /// <param name="rowIndex">row index</param>
        /// <returns>cell adress</returns>
        public static string ToCellAdress(int columnIndex, int rowIndex)
        {
            if (columnIndex < 1) throw (new ArgumentException("Invalid Argument. columnIndex must be > 0", "columnIndex"));
            if (rowIndex < 1) throw (new ArgumentException("Invalid Argument. rowIndex must be > 0", "rowIndex"));

            string returnValue = "";

            if (columnIndex <= 26)
            {
                returnValue = _columnIndex[columnIndex - 1] + rowIndex.ToString();
                return returnValue;
            }

            string preChar = "";
            int multi = columnIndex / _columnIndex.Length;
            preChar = _columnIndex[multi - 1];
            int columnArrayIndex = columnIndex;
            columnArrayIndex -= (multi * 26);
            returnValue = preChar + _columnIndex[columnArrayIndex - 1] + rowIndex.ToString();

            return returnValue;
        }

        /// <summary>
        /// Translate a color to double
        /// </summary>
        /// <param name="color">expression to convert</param>
        /// <returns>color</returns>
        public static double ToDouble(System.Drawing.Color color)
        {
            uint returnValue = color.B;
            returnValue = returnValue << 8;
            returnValue += color.G;
            returnValue = returnValue << 8;
            returnValue += color.R;
            return returnValue;
        }

        /// <summary>
        /// Translate a double to color
        /// </summary>
        /// <param name="color">expression to convert</param>
        /// <returns>color</returns>
        public static System.Drawing.Color ToColor(double color)
        {
            return System.Drawing.Color.FromArgb((int)color);
        }

        /// <summary>
        /// returns the valid file extension for the instance. for example ".xls" or ".xlsx"
        /// </summary>
        /// <param name="application">the instance</param>
        /// <returns>the extension</returns>
        public static string GetDefaultExtension(XlApplication application)
        { 
            double Version = Convert.ToDouble(application.Version);
            if (Version >= 120.00)
                return ".xlsx";
            else
                return ".xls";
        }
    }
}
