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
                for (var i = 0; i < args.Length; i++)
                {
                    var arg = args[i];

                    if (arg == "-d")
                    {
                        i = HandleDebugArguments(vs, args, i + 1);
                        continue;
                    }

                    if (arg.Contains('*') || arg.Contains('?'))
                    {
                        foreach (var file in GetFiles(arg))
                            vs.OpenFile(file);
                        continue;
                    }

                    vs.OpenFile(new FileInfo(arg));
                }
            }
            catch (Exception err)
            {
                Console.Error.WriteLine(err.ToString());
            }
        }

        static FileInfo[] GetFiles(string arg)
        {
            var directoryName = Path.GetDirectoryName(arg);
            var dir = new DirectoryInfo(string.IsNullOrEmpty(directoryName) ? Directory.GetCurrentDirectory() : directoryName);
            return dir.GetFiles(Path.GetFileName(arg));
        }

        static int HandleDebugArguments(VisualStudio vs, string[] args, int i)
        {
            if (i >= args.Length)
            {
                Console.Out.WriteLine(vs.DebugStartProgram);
                Console.Out.WriteLine(vs.DebugStartArguments);
                return i;
            }
            vs.DebugStartArguments = args[i];
            return i;
        }
    }
}
