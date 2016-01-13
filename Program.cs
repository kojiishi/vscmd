using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace vscmd {
    class Program {
        static VisualStudio vs;

        [STAThread]
        static void Main(string[] args) {
            try {
                vs = VisualStudio.GetOrCreate();
                bool activate = false;
                for (var i = 0; i < args.Length; i++) {
                    var arg = args[i];
                    if ("args".StartsWith(arg)) {
                        HandleDebugArguments(args.Skip(1));
                    } else if ("attach".StartsWith(arg)) {
                        HandleAttach(args.Skip(1));
                    } else if ("start".StartsWith(arg)) {
                        HandleDebugStart(args.Skip(1));
                    } else {
                        activate |= HandleFileArguments(args);
                    }
                    break;
                }
                if (activate)
                    vs.ActivateMainWindow();
            } catch (Exception err) {
                Console.Error.WriteLine(err.ToString());
            }
        }

        static FileInfo[] GetFiles(string arg) {
            var directoryName = Path.GetDirectoryName(arg);
            var dir = new DirectoryInfo(string.IsNullOrEmpty(directoryName) ? Directory.GetCurrentDirectory() : directoryName);
            return dir.GetFiles(Path.GetFileName(arg));
        }

        static bool HandleFileArguments(IEnumerable<string> args) {
            bool handled = false;
            foreach (var arg in args) {
                if (arg.Contains('*') || arg.Contains('?')) {
                    foreach (var file in GetFiles(arg)) {
                        vs.OpenFile(file);
                        handled = true;
                    }
                    continue;
                }

                vs.OpenFile(new FileInfo(arg));
                handled = true;
            }
            return handled;
        }

        static void HandleDebugStart(IEnumerable<string> args) {
            var config = vs.StartupProject.ActiveConfiguration;
            var program = args.FirstOrDefault();
            if (program == null) {
                Console.Out.WriteLine(config.DebugStartProgram);
                Console.Out.WriteLine(config.DebugStartArguments);
                return;
            }

            config.DebugStartProgram = Path.GetFullPath(program);

            args = args.Skip(1);
            if (args.Any())
                HandleDebugArguments(args);
        }

        static void HandleDebugArguments(IEnumerable<string> args) {
            var config = vs.StartupProject.ActiveConfiguration;
            if (!args.Any()) {
                Console.Out.WriteLine(config.DebugStartArguments);
                return;
            }

            args = args.Select(arg =>
                QuoteIfNeeded(File.Exists(arg) ? Path.GetFullPath(arg) : arg));
            config.DebugStartArguments = string.Join(" ", args);
        }

        static void HandleAttach(IEnumerable<string> args) {
            var filters = args
                .Select<string, Func<Process, bool>>(s => {
                    int id;
                    if (int.TryParse(s, out id))
                        return p => p.ProcessID == id;
                    return p => p.Name.Contains(s);
                }).ToList();
            var debugger = vs.Debugger;
            if (!filters.Any()) {
                foreach (var process in debugger.LocalProcesses)
                    Console.WriteLine($"{process.ProcessID} {process.Name}");
                return;
            }
            foreach (var process in debugger.LocalProcesses) {
                if (!filters.Any(f => f(process)))
                    continue;
                Console.WriteLine($"{process.ProcessID} {process.Name}");
                process.Attach();
            }
        }

        static string QuoteIfNeeded(string arg) {
            if (!arg.Contains(' '))
                return arg;
            return string.Concat("\"", arg, "\"");
        }
    }
}
