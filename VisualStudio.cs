using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace vscmd
{
    public class VisualStudio
    {
        private VisualStudio(dynamic dte)
        {
            this._dte = dte;
        }

        readonly dynamic _dte;

        public static VisualStudio GetOrCreate()
        {
            const string progID = "VisualStudio.DTE.12.0";
            dynamic dte;
            try
            {
                dte = Marshal.GetActiveObject(progID);
                dte.MainWindow.Activate();
            }
            catch (COMException err)
            {
                switch ((uint)err.HResult)
                {
                    case 0x800401E3: // MK_E_UNAVAILABLE: Operation unavailable
                        var type = Type.GetTypeFromProgID(progID);
                        dte = Activator.CreateInstance(type);
                        dte.UserControl = true;
                        break;
                    default:
                        throw;
                }
            }
            return new VisualStudio(dte);
        }

        public void OpenFile(string path)
        {
            path = Path.GetFullPath(path);
            this._dte.ExecuteCommand("File.OpenFile", "\"" + path + "\"");
        }
    }
}
