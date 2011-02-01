﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace LateBindingApi.CodeGenerator.Core.CodeGeneration
{
    static class EnumGenerator
    {
        #region Fields

        internal static string classHeader = "//Generated by LateBindingApi.CodeGenerator\r\nusing LateBindingApi.Core;\r\nnamespace %namespace%\r\n{\r\n";
        internal static string classFooter = "}";
        
        #endregion

        #region Internal Methods

        internal static string ConvertEnumToString(string nameSpace, XmlNode enumNode)
        {
            string result = classHeader.Replace("%namespace%", nameSpace);
            string attributes = Generator.GetVersionSummary(enumNode);
            string name = enumNode.Attributes["Name"].InnerText;
            result += attributes;
            result += "\tpublic enum " + name + Environment.NewLine + "\t{" + Environment.NewLine;

            XmlNode membersNode = enumNode.SelectSingleNode("Members");

            for (int i = 0; i < membersNode.ChildNodes.Count; i++)
            {
                XmlNode itemMember = membersNode.ChildNodes[i];

                string attributeLine = "\t" + Generator.GetVersionSummary(itemMember);
                string line = "\t\t" + itemMember.Attributes["Name"].InnerText + " = " + itemMember.Attributes["Value"].InnerText;

                result += attributeLine + line;

                if( i<(membersNode.ChildNodes.Count-1))
                    result += ",";

                result += Environment.NewLine;
            }

            result += "\t}" + Environment.NewLine;
            result += classFooter;
            return result;
        }

        internal static string ConvertEnumToFile(string directory, string nameSpace, XmlNode enumNode)
        {
            string fileName = System.IO.Path.Combine(directory, enumNode.Attributes["Name"].InnerText + ".cs");
            string newEnum = ConvertEnumToString(nameSpace, enumNode);
            System.IO.File.AppendAllText(fileName, newEnum);

            int i = directory.LastIndexOf("\\");
            string result = "\t\t<Compile Include=\""+ directory.Substring(i + 1) + "\\" + enumNode.Attributes["Name"].InnerText + ".cs" + "\" />";

            return result;
        }

        internal static string ConvertEnumsToFiles(string directory, string nameSpace, XmlNode enumsNode)
        {
            directory = System.IO.Path.Combine(directory, "Enums");
            if (false == System.IO.Directory.Exists(directory))
                System.IO.Directory.CreateDirectory(directory);

            nameSpace += ".Enums";

            string result="";
            foreach (XmlNode itemNode in enumsNode)
            {
                result += ConvertEnumToFile(directory, nameSpace, itemNode) + "\r\n"; 
            }
            return result;
        }

        #endregion
    }
}
