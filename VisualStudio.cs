using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace vscmd {
    public partial class VisualStudio : ComObject {
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
    }
}
