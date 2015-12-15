using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace vscmd
{
    public class VisualStudio
    {
        private VisualStudio(dynamic dte)
        {
            this._dte = dte;
        }

        const string ProgID = "VisualStudio.DTE";
        readonly dynamic _dte;

        enum HResult : uint
        {
            MK_E_UNAVAILABLE = 0x800401E3, // Operation unavailable
        }

        public static VisualStudio GetOrCreate()
        {
            dynamic dte = GetActiveObjectOrDefault();
            if (dte != null)
            {
                dte.MainWindow.Activate();
                return new VisualStudio(dte);
            }

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

        public void OpenFile(FileInfo file)
        {
            this._dte.ExecuteCommand("File.OpenFile", "\"" + file.FullName + "\"");
        }

        dynamic Solution
        {
            get { return this._dte.Solution; }
        }

        public class Project
        {
            Project(dynamic project)
            {
                this._project = project;
            }

            readonly dynamic _project;
        }
        dynamic ProjectByName(string name)
        {
            foreach (var item in this.Solution.Projects)
            {
                if (name == item.Name)
                    return item;
            }
            return null;
        }
        string StartupProjectName
        {
            get { return this.Solution.Properties.Item("StartupProject").Value; }
        }

        dynamic StartupProject
        {
            get { return this.ProjectByName(this.StartupProjectName); }
        }

        dynamic ActiveConfiguration
        {
            get { return this.StartupProject.ConfigurationManager.ActiveConfiguration; }
        }

        public string DebugStartProgram
        {
            get
            {
                var config = this.ActiveConfiguration;
                int action = config.Properties.Item("StartAction").Value;
                return action == 0 ? null : config.Properties.Item("StartProgram").Value;
            }
            set
            {
                var config = this.ActiveConfiguration;
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
            }
        }

        public string DebugStartArguments
        {
            get { return this.ActiveConfiguration.Properties.Item("StartArguments").Value; }
            set { this.ActiveConfiguration.Properties.Item("StartArguments").Value = value; }
        }
    }
}
