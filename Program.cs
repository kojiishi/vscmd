using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace vscmd
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                OpenFiles(args);
            }
            catch (Exception err)
            {
                Console.Error.WriteLine(err.ToString());
            }
        }

        static void OpenFiles(string[] files)
        {
            dynamic dte = Marshal.GetActiveObject("VisualStudio.DTE.12.0");

            foreach (var item in files)
            {
                var path = Path.GetFullPath(item);
                dte.ExecuteCommand("File.OpenFile", "\"" + path + "\"");
            }
        }
    }
}
