﻿using System;
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
        static VisualStudio vs;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                vs = VisualStudio.GetOrCreate();
                bool activate = false;
                for (var i = 0; i < args.Length; i++)
                {
                    var arg = args[i];
                    if ("args".StartsWith(arg)) {
                        HandleDebugArguments(args.Skip(1));
                    } else {
                        activate |= HandleFileArguments(args);
                    }
                    break;
                }
                if (activate)
                    vs.ActivateMainWindow();
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

        static bool HandleFileArguments(IEnumerable<string> args)
        {
            bool handled = false;
            foreach (var arg in args)
            {
                if (arg.Contains('*') || arg.Contains('?'))
                {
                    foreach (var file in GetFiles(arg))
                    {
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

        static void HandleDebugArguments(IEnumerable<string> args)
        {
            var project = vs.StartupProject;
            var config = project.ActiveConfiguration;
            if (!args.Any())
            {
                Console.Out.WriteLine(config.DebugStartArguments);
                return;
            }

            args = args.Select(arg => File.Exists(arg) ? Path.GetFullPath(arg) : arg);
            config.DebugStartArguments = string.Join(" ", args);
        }
    }
}
