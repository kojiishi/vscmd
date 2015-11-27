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
                var vs = VisualStudio.GetOrCreate();
                foreach (var arg in args)
                {
                    var directoryName = Path.GetDirectoryName(arg);
                    var dir = new DirectoryInfo(string.IsNullOrEmpty(directoryName) ? Directory.GetCurrentDirectory() : directoryName);
                    foreach (var file in dir.GetFiles(Path.GetFileName(arg)))
                    {
                        vs.OpenFile(file);
                    }
                }
            }
            catch (Exception err)
            {
                Console.Error.WriteLine(err.ToString());
            }
        }
    }
}
