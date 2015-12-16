using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace vscmd
{
    public class ComObject
    {
        public ComObject(object obj)
        {
            this.Object = obj;
        }

        protected dynamic Object { get; }
    }

    public class VisualStudio : ComObject
    {
        #region Constructors

        private VisualStudio(object dte)
            : base(dte)
        {
        }

        const string ProgID = "VisualStudio.DTE";

        enum HResult : uint
        {
            MK_E_UNAVAILABLE = 0x800401E3, // Operation unavailable
        }

        public static VisualStudio GetOrCreate()
        {
            dynamic dte = GetActiveObjectOrDefault();
            if (dte != null)
                return new VisualStudio(dte);

            var type = Type.GetTypeFromProgID(ProgID);
            dte = Activator.CreateInstance(type);
            dte.UserControl = true;

            return new VisualStudio(dte);
        }

        static object GetActiveObjectOrDefault()
        {
            foreach (var progID in GetProgIDs())
            {
                try
                {
                    return Marshal.GetActiveObject(progID);
                }
                catch (COMException err)
                {
                    if ((HResult)err.HResult != HResult.MK_E_UNAVAILABLE)
                        throw;
                }
            }
            return null;
        }

        static IEnumerable<string> GetProgIDs()
        {
            yield return ProgID;

            const string prefix = ProgID + ".";
            var versioned = Registry.ClassesRoot.GetSubKeyNames()
                .Where(progID => progID.StartsWith(prefix))
                .OrderByDescending(progID => float.Parse(progID.Substring(prefix.Length)));
            foreach (var progID in versioned)
                yield return progID;
        }

        #endregion

        #region Window

        public void ActivateMainWindow()
        {
            this.Object.MainWindow.Activate();
        }

        #endregion

        #region ExecuteCommand

        public void OpenFile(FileInfo file)
        {
            this.Object.ExecuteCommand("File.OpenFile", "\"" + file.FullName + "\"");
        }

        #endregion

        #region Solution

        dynamic Solution
        {
            get { return this.Object.Solution; }
        }

        #endregion

        #region Project

        IEnumerable<dynamic> ProjectObjectsRecursive()
        {
            foreach (var item in this.Solution.Projects)
            {
                foreach (var project in ProjectObjectsFromItem(item))
                    yield return project;
            }
        }

        static IEnumerable<dynamic> ProjectObjectsFromItem(dynamic item, int level = 0)
        {
            //Console.WriteLine(new string(' ', level * 4) + item.Name + "=" + item.Kind);
            if (item.ConfigurationManager != null)
            {
                yield return item;
                yield break;
            }

            var projectItems = item.ProjectItems;
            if (projectItems != null)
            {
                foreach (var projectItem in projectItems)
                {
                    var subProjectItem = projectItem.SubProject;
                    if (subProjectItem != null)
                    {
                        foreach (var subProject in ProjectObjectsFromItem(subProjectItem, level + 1))
                            yield return subProject;
                    }
                }
            }
        }

        Project ProjectByName(string name)
        {
            foreach (var item in this.ProjectObjectsRecursive())
            {
                if (name == item.Name)
                    return new Project(item);
            }
            throw new KeyNotFoundException(name);
        }
        string StartupProjectName
        {
            get { return this.Solution.Properties.Item("StartupProject").Value; }
        }

        public Project StartupProject
        {
            get { return this.ProjectByName(this.StartupProjectName); }
        }

        public enum ProjectKind
        {
            Unknown,
            CPlusPlus,
            CSharp,
        }

        public class Project : ComObject
        {
            internal Project(object project)
                : base(project)
            {
                this.Kind = ParseKind(project);
            }

            public string Name { get { return this.Object.Name; } }
            public ProjectKind Kind { get; }

            static ProjectKind ParseKind(dynamic project)
            {
                switch ((string)project.Kind)
                {
                    case "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}": return ProjectKind.CPlusPlus;
                    case "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}": return ProjectKind.CSharp;
                    default: return ProjectKind.Unknown;
                }
            }

            NotSupportedException KindNotSupportedException()
            {
                return new NotSupportedException(string.Format(
                    "The project type {0} is not supported", this.Kind));
            }

            dynamic ActiveConfiguration
            {
                get { return this.Object.ConfigurationManager.ActiveConfiguration; }
            }

            public string DebugStartProgram
            {
                get
                {
                    var config = this.ActiveConfiguration;
                    switch (this.Kind)
                    {
                        case ProjectKind.CPlusPlus:
                            return config.Properties.Item("Command").Value;
                        case ProjectKind.CSharp:
                            int action = config.Properties.Item("StartAction").Value;
                            return action == 0 ? null : config.Properties.Item("StartProgram").Value;
                        default:
                            throw this.KindNotSupportedException();
                    }
                }
                set
                {
                    var config = this.ActiveConfiguration;
                    switch (this.Kind)
                    {
                        case ProjectKind.CPlusPlus:
                            config.Properties.Item("Command").Value = value;
                            break;
                        case ProjectKind.CSharp:
                            if (string.IsNullOrEmpty(value))
                            {
                                config.Properties.Item("StartAction").Value = 0;
                                return;
                            }
                            if (File.Exists(value))
                            {
                                config.Properties.Item("StartAction").Value = 1;
                                config.Properties.Item("StartProgram").Value = new FileInfo(value).FullName;
                                return;
                            }
                            config.Properties.Item("StartAction").Value = 1;
                            config.Properties.Item("StartProgram").Value = value;
                            break;
                        default:
                            throw this.KindNotSupportedException();
                    }
                }
            }

            public string DebugStartArguments
            {
                get { return this.ActiveConfiguration.Properties.Item(this.DebugStartArgumentsPropertyName).Value; }
                set { this.ActiveConfiguration.Properties.Item(this.DebugStartArgumentsPropertyName).Value = value; }
            }

            string DebugStartArgumentsPropertyName
            {
                get
                {
                    switch (this.Kind)
                    {
                        case ProjectKind.CPlusPlus: return "CommandArguments";
                        case ProjectKind.CSharp: return "StartArguments";
                        default: throw this.KindNotSupportedException();
                    }
                }
            }

        }

        #endregion
    }
}
