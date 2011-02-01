using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;

using LateBindingApi.CodeGenerator.Core;

namespace LateBindingApi.CodeGenerator.Core.CodeGeneration
{
    public static class Generator
    {
        #region Fields

        internal static string _methodeInvokeCall;
        internal static string _methodeInvokeCallWithScalarReturn;
        internal static string _methodeInvokeCallWithComReturn;
        
        #endregion

        #region Properties
        
        public static string MethodInvokeCall
        {
            get 
            {
                if (null == _methodeInvokeCall)
                {
                    _methodeInvokeCall = ReadTextFileFromRessource("Ressources.MethodInvokeCall.txt"); 
                }
                return _methodeInvokeCall;
            }
            set 
            {
                _methodeInvokeCall = value;
            }
        }

        public static string MethodInvokeCallWithScalarReturn
        {
            get
            {
                if (null == _methodeInvokeCallWithScalarReturn)
                {
                    _methodeInvokeCallWithScalarReturn = ReadTextFileFromRessource("Ressources.MethodInvokeCallWithReturn.txt");
                }
                return _methodeInvokeCallWithScalarReturn;
            }
            set
            {
                _methodeInvokeCallWithScalarReturn = value;
            }
        }

        public static string MethodInvokeCallWithComReturn
        {
            get
            {
                if (null == _methodeInvokeCallWithComReturn)
                {
                    _methodeInvokeCallWithComReturn = ReadTextFileFromRessource("Ressources.MethodInvokeCallWithComReturn.txt");
                }
                return _methodeInvokeCallWithComReturn;
            }
            set
            {
                _methodeInvokeCallWithComReturn = value;
            }
        }
        
        #endregion

        private static string GetProjectEntryInSolution(string projectName)
        {
            string returnValue = "Project(\"{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}\") = \"%ProjectName%\", \"%ProjectName%\\%ProjectName%.csproj\", \"{94BC6D99-3FD2-4520-B8D5-E7F6E9739859}\"\r\nEndProject\r\n";
            returnValue = returnValue.Replace("%ProjectName%", projectName);
            return returnValue;
        }

        private static string CreateProjectDirectory(string rootPath, string projectName)
        {
            string projectDirectory = Path.Combine(rootPath, projectName);

            if (Directory.Exists(projectDirectory) == false)
                Directory.CreateDirectory(projectDirectory);

            return projectDirectory;
        }

        private static string CreateProjectFile(string projectName)
        {
            string projectFile = ReadTextFileFromRessource("VS2008ProjectTemplate.Project.csproj");
            projectFile = projectFile.Replace("%ProjectName%", projectName);
            return projectFile;
        }

        public static void CreateSampleSolution(string directory)
        {
            COMComponentReader reader = new COMComponentReader();
            string innerXml = ReadTextFileFromRessource("LateBindingApi.CodeGenerator.SampleDocument.xml");
            reader.COMTree.LoadXml(innerXml);
            CreateSolution(reader, directory);
        }

        public static void CreateSolution(COMComponentReader reader, string directory)
        {
            XmlNode solutionNode = reader.COMTree.SelectSingleNode(Constants.Solution);
            string solutionName = solutionNode.Attributes["Name"].InnerText;
            string solutionPrefix = solutionNode.Attributes["Prefix"].InnerText;

            #region Directory

            string rootPath = Path.Combine(directory, solutionName);
            if (Directory.Exists(directory) == false)
                Directory.CreateDirectory(directory);

            if (Directory.Exists(rootPath) == true)
                Directory.Delete(rootPath, true);
            Directory.CreateDirectory(rootPath);

            #endregion

            string solutionFile = ReadTextFileFromRessource("Solution.sln");
            solutionFile = solutionFile.Replace("%SolutionName%", solutionName);

            string projectString = "";
            foreach (XmlNode projectNode in solutionNode.ChildNodes)
            {
                #region Project

                string projectName = projectNode.Attributes["ProjectName"].InnerText;
                if (false == Directory.Exists(Path.Combine(rootPath, projectName)))
                    Directory.CreateDirectory(Path.Combine(rootPath, projectName));
                projectString += GetProjectEntryInSolution(projectName);
                string projectDirectory =  CreateProjectDirectory(rootPath, projectName);
                string projectFile = CreateProjectFile(projectName);
                CreateFile("VS2008ProjectTemplate.AssemblyInfo.cs", Path.Combine(projectDirectory, "AssemblyInfo.cs"), projectName);

                #endregion

                XmlNode enumsNode = projectNode.SelectSingleNode("Enums");
                string enumsInclude = EnumGenerator.ConvertEnumsToFiles(projectDirectory, projectName, enumsNode);
                projectFile = projectFile.Replace("%EnumIncludes%", enumsInclude);

                XmlNode dispatchesNode = projectNode.SelectSingleNode("DispatchInterfaces");
                string dispatchesInclude = DispatchGenerator.ConvertDispatchesToFiles(projectDirectory, projectName, solutionPrefix, dispatchesNode);
                projectFile = projectFile.Replace("%DispatchIncludes%", dispatchesInclude);
              
                File.AppendAllText(Path.Combine(rootPath, projectName + "\\" + projectName + ".csproj"), projectFile, Encoding.UTF8);

            }
            solutionFile = solutionFile.Replace("%Projects%",projectString);

            byte[] coreBinary = ReadBinaryFromResource("LateBindingApi.dll");
            WriteBinaryToFile(coreBinary, Path.Combine(rootPath, "LateBindingApi.dll"));

            File.AppendAllText(Path.Combine(rootPath,solutionName + ".sln"), solutionFile, Encoding.UTF8);
        }

        public static void CreateCoreSolution(string directory, string projectName)
        {
            #region Directory

            string rootPath = Path.Combine(directory, projectName);
            if (Directory.Exists(directory) == false)
                Directory.CreateDirectory(directory);

            if (Directory.Exists(rootPath) == true )
                Directory.Delete(rootPath,true);
            Directory.CreateDirectory(rootPath);

            string sourceFolderPath = Path.Combine(rootPath, "SourceFiles");
            if (Directory.Exists(sourceFolderPath) == false)
                Directory.CreateDirectory(sourceFolderPath);

            string interfacePath = Path.Combine(sourceFolderPath, "Interfaces");
            if (Directory.Exists(interfacePath) == false)
                Directory.CreateDirectory(interfacePath);
            
            #endregion

            #region Create Files

            CreateFile("VS2008CoreTemplate.Solution.sln", Path.Combine(rootPath, projectName + ".sln"), projectName);
            CreateFile("VS2008CoreTemplate.Project.csproj", Path.Combine(sourceFolderPath, projectName + ".csproj"), projectName);
            CreateFile("VS2008CoreTemplate.AssemblyInfo.cs", Path.Combine(sourceFolderPath, "AssemblyInfo.cs"), projectName);
            CreateFile("VS2008CoreTemplate.SupportByLibraryAttribute.cs", Path.Combine(sourceFolderPath, "SupportByLibraryAttribute.cs"), projectName);
            CreateFile("VS2008CoreTemplate.Settings.cs", Path.Combine(sourceFolderPath, "Settings.cs"), projectName);

            CreateFile("VS2008CoreTemplate.Interfaces.IObject.cs", Path.Combine(interfacePath, "IObject.cs"), projectName);
            CreateFile("VS2008CoreTemplate.Interfaces.IEventBinding.cs", Path.Combine(interfacePath, "IEventBinding.cs"), projectName);
            CreateFile("VS2008CoreTemplate.Interfaces.INonCreatable.cs", Path.Combine(interfacePath, "INonCreatable.cs"), projectName);
            CreateFile("VS2008CoreTemplate.Interfaces.ICreatable.cs", Path.Combine(interfacePath, "ICreatable.cs"), projectName);

            CreateFile("VS2008CoreTemplate.COMObject.cs", Path.Combine(sourceFolderPath, "COMObject.cs"), projectName);
            CreateFile("VS2008CoreTemplate.CreatableObject.cs", Path.Combine(sourceFolderPath, "CreatableObject.cs"), projectName);
            CreateFile("VS2008CoreTemplate.NonCreatableObject.cs", Path.Combine(sourceFolderPath, "NonCreatableObject.cs"), projectName);
            CreateFile("VS2008CoreTemplate.Invoker.cs", Path.Combine(sourceFolderPath, "Invoker.cs"), projectName);
            CreateFile("VS2008CoreTemplate.SinkHelper.cs", Path.Combine(sourceFolderPath, "SinkHelper.cs"), projectName);
            
            #endregion
        }

        #region Methods

        internal static string ReadTextFileFromRessource(string fileName)
        {
            fileName = "LateBindingApi.CodeGenerator.Core.CodeGeneration." + fileName;

            System.IO.Stream ressourceStream;
            System.IO.StreamReader textStreamReader;
            try
            {
                ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(fileName);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                ressourceStream.Close();
                textStreamReader.Close();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }

        }

        internal static void WriteBinaryToFile(byte[] binary, string fileName)
        {
            System.IO.FileStream fs = new System.IO.FileStream(fileName, System.IO.FileMode.Create);
            fs.Write(binary, 0, binary.Length);
            fs.Close();
        }

        internal static byte[] ReadBinaryFromResource(string resourceName)
        {
            resourceName = "LateBindingApi.CodeGenerator.Core.CodeGeneration." + resourceName;

            System.IO.Stream ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
            byte[] binary = new byte[ressourceStream.Length];
            ressourceStream.Read(binary, 0, binary.Length);
            ressourceStream.Close();
            return binary;
        }
        
        internal static void CreateFile(string ressourceName, string fileName, string projectName)
        {
            string projectFile = ReadTextFileFromRessource(ressourceName);
            projectFile = projectFile.Replace("%ProjectName%", projectName);
            File.AppendAllText(fileName, projectFile, Encoding.UTF8);
        }

        internal static string GetVersionSummary(XmlNode enumNode)
        {
            string result = "";
            string versions = "";
            XmlNode components = enumNode.SelectSingleNode("Components");
            foreach (XmlNode itemNode in components.ChildNodes)
            {
                string key = itemNode.Name;
                XmlNode rootComponentNode = GetComponentNode(itemNode);
                string versionAttribute = GetChildInnerText(rootComponentNode, "Version");
                versions += "\"" + versionAttribute + "\"";                
                if (itemNode != components.LastChild)
                    versions += ",";
            }

            result += "\t[SupportByLibrary(" + versions + ")]" + Environment.NewLine;
            return result;
        }

        internal static string GetChildInnerText(XmlNode node, string name)
        {
            foreach (XmlNode item in node.ChildNodes)
            {
                if (item.Name == name)
                    return item.InnerText;
            }
            throw (new ArgumentException("Node not found" + name));
        }

        internal static XmlNode GetComponentNode(XmlNode refComponentNode)
        {
            string key = refComponentNode.InnerText;
            XmlNode rootComponentsNode = refComponentNode.OwnerDocument.SelectSingleNode(Constants.Components);

            foreach (XmlNode childNode in rootComponentsNode.ChildNodes)
            {
                XmlNode keyNode = GetChildNode(childNode, "Key");
                if (key == keyNode.InnerText)
                    return childNode;
            }

            throw (new ArgumentException("Node not found. " + key));
        }

        internal static XmlNode GetChildNode(XmlNode node, string name)
        {
            foreach (XmlNode item in node.ChildNodes)
            {
                if (item.Name == name)
                    return item;
            }
            throw (new ArgumentException("Node not found" + name));
        }

        #endregion
    }
}
