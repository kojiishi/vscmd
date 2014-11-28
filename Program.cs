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
                    vs.OpenFile(arg);
            }
            catch (Exception err)
            {
                Console.Error.WriteLine(err.ToString());
            }
        }
    }
}
