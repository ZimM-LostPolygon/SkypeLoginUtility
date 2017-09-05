using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace SkypeConsoleLogin {
    internal class Program {
        [STAThread]
        private static void Main(string[] args) {
            System.Diagnostics.Process.GetProcessesByName("skype").ToList().ForEach(process => process.Kill());

            NativeMethods.AttachConsole(-1);

            string username = null;
            string password = null;
            string skypeExePath = null;

            string skypeArguments = "";
            for (int i = 0; i < args.Length; i++) {
                string arg = args[i];
                if (i == 0) {
                    skypeExePath = arg;
                    continue;
                }

                Match matchUsername = Regex.Match(arg, "/username:(.+)", RegexOptions.IgnoreCase);
                if (matchUsername.Success) {
                    username = matchUsername.Groups[1].Value;
                    continue;
                }

                Match matchPassword = Regex.Match(arg, "/password:(.+)", RegexOptions.IgnoreCase);
                if (matchPassword.Success) {
                    password = matchPassword.Groups[1].Value;
                    continue;
                }

                skypeArguments += arg + " ";
            }

            if (!File.Exists(skypeExePath)) {
                Console.WriteLine("Skype executable not found at the provided path");
                return;
            }

            if (username == null || password == null) {
                Console.WriteLine("Usage: SkypeConsoleLogin.exe <path to Skype.exe> /username:<username> /password:<password> [other Skype arguments]");
                return;
            }

            //try {
                SkypeLauncher launcher = new SkypeLauncher(skypeExePath, username, password, skypeArguments);
                launcher.Launch();
            //} catch (Exception e) {
            //    Console.WriteLine("Error while starting Skype:\r\n" + e);
            //    if (Environment.UserInteractive) {
            //        Console.Read();
            //    }
            //}

            NativeMethods.FreeConsole();
        }

        private static class NativeMethods {
            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool AllocConsole();

            [DllImport("kernel32.dll", SetLastError = true)]
            public static extern bool FreeConsole();

            [DllImport("kernel32", SetLastError = true)]
            public static extern bool AttachConsole(int dwProcessId);
        }
    }
}