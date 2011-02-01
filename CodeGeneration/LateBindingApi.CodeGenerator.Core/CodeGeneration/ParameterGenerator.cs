using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.CodeGeneration
{
    static class ParameterGenerator
    {
        internal static bool HasParameters(XmlNode methodNode)
        { 
            if(methodNode.SelectSingleNode("Parameters").ChildNodes.Count > 0)
                return true;
            else
                return false;
                    
        }

        internal static string GetParamsArrayName(XmlNode itemNode)
        {
            if(true == HasParameters(itemNode))
                return "paramsArray";
            else
                return "null";
        }

        internal static string GetParamsArray(XmlNode itemNode)
        {
            XmlNode parametersNode = itemNode.SelectSingleNode("Parameters");
            if (0 == parametersNode.ChildNodes.Count)
                return "";

            string returnString = "\t\t\tobject[] paramsArray = new object[" + parametersNode.ChildNodes.Count + "];\r\n";

            int i = 0;
            foreach (XmlNode paramNode in parametersNode.ChildNodes)
            {
                string paramType = paramNode.Attributes["Type"].InnerText;
                string paramName = paramNode.Attributes["Name"].InnerText;
                paramName = FirstCaseLower(paramName);
                returnString += "\t\t\tparamsArray[" + i.ToString() + "] = " + paramName + ";\r\n";

                i++;
            }
            return returnString;
        }

        internal static string GetParamsString(XmlNode itemNode)
        {
            XmlNode parametersNode = itemNode.SelectSingleNode("Parameters");
            if (parametersNode.ChildNodes.Count == 0)
                return "()";

            string returnString = "(";

            foreach (XmlNode paramNode in parametersNode.ChildNodes)
            {
                string paramType = paramNode.Attributes["Type"].InnerText;
                string paramName = paramNode.Attributes["Name"].InnerText;
                paramName = FirstCaseLower(paramName);
                returnString += paramType + " " + paramName;

                if (paramNode != parametersNode.LastChild)
                    returnString += ", ";

            }
            return returnString + ")";
        }

        private static string FirstCaseLower(string expression)
        {
            if ((null == expression) || expression.Length < 1)
                return expression;

            string firstCase = expression.Substring(0, 1).ToLower();
            return firstCase + expression.Substring(1);           
        }
    }
}
