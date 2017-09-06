using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using LostPolygon.WinApiWrapper;

namespace SkypeConsoleLogin {
    internal class Program {
        [STAThread]
        private static void Main(string[] args) {
#if DEBUG
            Process.GetProcessesByName("skype").ToList().ForEach(process => process.Kill());
#endif
            WinApi.Kernel32.AttachConsole(-1);

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

#if !DEBUG
            try {
#endif
                SkypeLauncher launcher = new SkypeLauncher(skypeExePath, username, password, skypeArguments);
                launcher.Launch();
#if !DEBUG
            } catch (Exception e) {
                Console.WriteLine("Error while starting Skype:\r\n" + e);
                if (Environment.UserInteractive) {
                    Console.Read();
                }
            }
#endif

            WinApi.Kernel32.FreeConsole();
        }
    }
}
