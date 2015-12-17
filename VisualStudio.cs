using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace vscmd {
    public class VisualStudio : ComObject {
        #region Constructors

        private VisualStudio(object dte)
            : base(dte) { }

        const string ProgID = "VisualStudio.DTE";

        enum HResult : uint {
            MK_E_UNAVAILABLE = 0x800401E3, // Operation unavailable
        }

        public static VisualStudio GetOrCreate() {
            dynamic dte = GetActiveObjectOrDefault();
            if (dte != null)
                return new VisualStudio(dte);

            var type = Type.GetTypeFromProgID(ProgID);
            dte = Activator.CreateInstance(type);
            dte.UserControl = true;

            return new VisualStudio(dte);
        }

        static object GetActiveObjectOrDefault() {
            foreach (var progID in GetProgIDs()) {
                try {
                    return Marshal.GetActiveObject(progID);
                } catch (COMException err) {
                    if ((HResult)err.HResult != HResult.MK_E_UNAVAILABLE)
                        throw;
                }
            }
            return null;
        }

        static IEnumerable<string> GetProgIDs() {
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

        public void ActivateMainWindow() {
            this.Object.MainWindow.Activate();
        }

        #endregion

        #region ExecuteCommand

        public void OpenFile(FileInfo file) {
            this.Object.ExecuteCommand("File.OpenFile", "\"" + file.FullName + "\"");
        }

        #endregion

        #region Solution

        dynamic Solution {
            get { return this.Object.Solution; }
        }

        #endregion

        #region Project

        IEnumerable<dynamic> ProjectObjectsRecursive() {
            foreach (var item in this.Solution.Projects) {
                foreach (var project in ProjectObjectsFromItem(item))
                    yield return project;
            }
        }

        static IEnumerable<dynamic> ProjectObjectsFromItem(dynamic item, int level = 0) {
            //Console.WriteLine(new string(' ', level * 4) + item.Name + "=" + item.Kind);
            if (item.ConfigurationManager != null) {
                yield return item;
                yield break;
            }

            var projectItems = item.ProjectItems;
            if (projectItems != null) {
                foreach (var projectItem in projectItems) {
                    var subProjectItem = projectItem.SubProject;
                    if (subProjectItem != null) {
                        foreach (var subProject in ProjectObjectsFromItem(subProjectItem, level + 1))
                            yield return subProject;
                    }
                }
            }
        }

        Project ProjectByName(string name) {
            foreach (var item in this.ProjectObjectsRecursive()) {
                if (name == item.Name)
                    return Project.FromObject(item);
            }
            throw new KeyNotFoundException(name);
        }
        string StartupProjectName {
            get { return this.Solution.Properties.Item("StartupProject").Value; }
        }

        public Project StartupProject {
            get { return this.ProjectByName(this.StartupProjectName); }
        }

        public partial class Project : ComObject {
            protected Project(object project)
                : base(project) { }

            internal static Project FromObject(dynamic project) {
                switch ((string)project.Kind) {
                case "{8BC9CEB8-8B4A-11D0-8D11-00A0C91BC942}":
                    return new CPlusPlusProject(project);
                case "{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}":
                    return new CSharpProject(project);
                default:
                    return new Project(project);
                }
            }

            public string Name { get { return this.Object.Name; } }

            internal NotSupportedException KindNotSupportedException() {
                return new NotSupportedException(string.Format(
                    "The operation is not suuported for this project type {0}.", this.Object.Kind));
            }
        }

        partial class CPlusPlusProject : Project {
            internal CPlusPlusProject(object project)
                : base(project) { }
        }

        partial class CSharpProject : Project {
            internal CSharpProject(object project)
                : base(project) { }
        }

        #endregion

        #region Configuration

        partial class Project {
            public Configuration ActiveConfiguration {
                get { return this.CreateConfiguration(this.Object.ConfigurationManager.ActiveConfiguration); }
            }

            protected virtual Configuration CreateConfiguration(object configuration) { return new Configuration(configuration, this); }
        }

        public class Configuration : ComChildObject<Project> {
            internal Configuration(object configuration, Project parent)
                : base(configuration, parent) { }

            public Project Project { get { return this.Parent; } }
            public string Name { get { return this.Object.ConfigurationName; } }

            public virtual string DebugStartArguments {
                get { throw this.Project.KindNotSupportedException(); }
                set { throw this.Project.KindNotSupportedException(); }
            }

            public virtual string DebugStartProgram {
                get { throw this.Project.KindNotSupportedException(); }
                set { throw this.Project.KindNotSupportedException(); }
            }
        }

        partial class CPlusPlusProject {
            protected override Configuration CreateConfiguration(object configuration) { return new CPlusPlusConfiguration(configuration, this); }
        }

        class CPlusPlusConfiguration : Configuration {
            public CPlusPlusConfiguration(object configuration, Project parent)
                : base(configuration, parent) { }

            public override string DebugStartArguments {
                get { return this.PropertyValue<string>("CommandArguments"); }
                set { this.SetPropertyValue<string>("CommandArguments", value); }
            }

            public override string DebugStartProgram {
                get { return this.PropertyValue<string>("Command"); }
                set { this.SetPropertyValue("Command", value); }
            }
        }

        partial class CSharpProject {
            protected override Configuration CreateConfiguration(object configuration) { return new CSharpConfiguration(configuration, this); }
        }

        class CSharpConfiguration : Configuration {
            public CSharpConfiguration(object configuration, Project parent)
                : base(configuration, parent) { }

            public override string DebugStartArguments {
                get { return this.PropertyValue<string>("StartArguments"); }
                set { this.SetPropertyValue<string>("StartArguments", value); }
            }

            public override string DebugStartProgram {
                get {
                    int action = PropertyValue<int>("StartAction");
                    return action == 0 ? null : PropertyValue<string>("StartProgram");
                }

                set {
                    if (string.IsNullOrEmpty(value)) {
                        SetPropertyValue("StartAction", 0);
                        return;
                    }
                    if (File.Exists(value)) {
                        SetPropertyValue("StartAction", 1);
                        SetPropertyValue("StartProgram", new FileInfo(value).FullName);
                        return;
                    }
                    SetPropertyValue("StartAction", 1);
                    SetPropertyValue("StartProgram", value);
                }
            }
        }

        #endregion
    }
}
