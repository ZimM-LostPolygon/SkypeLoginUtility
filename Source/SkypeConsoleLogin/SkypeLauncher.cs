using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using LostPolygon.WinApiWrapper;
using mshtml;

namespace SkypeConsoleLogin {
    internal class SkypeLauncher {
        private const int kOperationRetryDelay = 30;
        private const int kSkypeRestartDelay = 1000;
        private const int kSkypeMaxRestarts = 3;
        private const int kWaitForSkypeLoginWindowTimeout = 12000;
        private const int kHtmlGetObjectTimeout = 3000;
        private const int kLoginPageReadyTimeout = 10000;
        private readonly string _skypeExePath;
        private readonly string _arguments;
        private readonly string _username;
        private readonly string _password;
        private readonly bool _minimized;



        private Process _skypeProcess;
        private HTMLDocumentClass _loginBrowserHtmlDocument;
        private volatile bool _canThrowSkypeExitException;
        private Thread _currentThread;

        public Process SkypeProcess {
            get { return _skypeProcess; }
        }

        public SkypeLauncher(string skypeExePath, string username, string password, string arguments) {
            _skypeExePath = skypeExePath;
            _username = username;
            _password = password;
            _arguments = arguments;

            _minimized = arguments.Split(' ').Any(s => s.ToLowerInvariant() == "/minimized");
        }

        public void Launch() {
            _currentThread = Thread.CurrentThread;
            StartSkypeProcess();

            Thread skypeWaitForExitThread = new Thread(() => {
                _skypeProcess.WaitForExit();
                if (_canThrowSkypeExitException) {
                    _currentThread.Abort();
                }
            }) { IsBackground = true };
            skypeWaitForExitThread.Start();

            try {
                ExecuteLogin();
            } finally {
                if (skypeWaitForExitThread.IsAlive) {
                    skypeWaitForExitThread.Abort();
                }
            }
        }

        private void ExecuteLogin() {
            // Wait for Skype login window to create
            Debug.WriteLine("Wait for Skype login window to create");

            GetSkypeLoginWindowDataResult skypeLoginWindowData = new GetSkypeLoginWindowDataResult(GetSkypeLoginWindowDataResult.State.Failed);
            int skypeRestarts = 0;
            bool success = false;
            while (skypeRestarts < kSkypeMaxRestarts) {
                bool isSkypeExited = false;
                bool isAlreadyLoggedIn = false;
                success = TimedOutOperation(kWaitForSkypeLoginWindowTimeout, kOperationRetryDelay, () => {
                    if (_skypeProcess.HasExited) {
                        isSkypeExited = true;

                        // Exit prematurely
                        return true;
                    }

                    skypeLoginWindowData = GetSkypeLoginWindowData();
                    switch (skypeLoginWindowData.LoginState) {
                        case GetSkypeLoginWindowDataResult.State.Success:
                            return true;
                        case GetSkypeLoginWindowDataResult.State.Failed:
                            return false;
                        case GetSkypeLoginWindowDataResult.State.AlreadyLoggedIn:
                            isAlreadyLoggedIn = true;
                            return true;
                        default:
                            throw new ArgumentOutOfRangeException();
                    }
                });

                if (isAlreadyLoggedIn) {
                    Console.WriteLine("Already logged in");
                    return;
                }

                if (isSkypeExited) {
                    // Restarts Skype after some delay
                    success = false;
                    skypeRestarts++;

                    Console.WriteLine("Skype process exited unexpectedly, restarting in {0} ms", kSkypeRestartDelay);
                    Thread.Sleep(kSkypeRestartDelay);
                    StartSkypeProcess();
                } else {
                    break;
                }
            }

            if (!success)
                throw new LoginException("Unable to detect Skype login window");

            Login(skypeLoginWindowData);
        }

        private void Login(GetSkypeLoginWindowDataResult skypeLoginWindowData) {
            // Set user name
            WinApi.User32.SendMessage(skypeLoginWindowData.LoginEditPtr, WinApi.MessageType.WM_SETTEXT, IntPtr.Zero, _username);

            // Simulate Enter press to initiate login
            WinApi.User32.PostMessage(skypeLoginWindowData.LoginEditPtr, WinApi.MessageType.WM_KEYDOWN, new IntPtr((int) WinApi.VirtualKeyCode.VK_RETURN), IntPtr.Zero);

            // Force minimize Skype if it should. For some reason Skype restores itself after login begins
            if (_minimized) {
                TimedOutOperation(1000, 1, () => {
                    if (!WinApi.User32.IsIconic(skypeLoginWindowData.LoginWindowPtr)) {
                        WinApi.User32.ShowWindow(skypeLoginWindowData.LoginWindowPtr, WinApi.User32.ShowWindowCommands.SW_FORCEMINIMIZE);
                        return true;
                    }
                    return false;
                });
            }

            _canThrowSkypeExitException = true;
            try {
                bool success = TimedOutOperation(kHtmlGetObjectTimeout, kOperationRetryDelay, () => {
                    IntPtr internetExplorerServerHandle = GetInternetExplorerServerHandle(skypeLoginWindowData.LoginWindowPtr);
                    if (internetExplorerServerHandle == IntPtr.Zero)
                        return false;

                    _loginBrowserHtmlDocument = GetHtmlDocumentClassFromInternetExplorerServerHandle(internetExplorerServerHandle);
                    if (_loginBrowserHtmlDocument == null)
                        return false;

                    return true;
                });

                if (!success)
                    throw new LoginException("Unable to get the Skype login web document");

                WebBrowserLogin();
            } catch (ThreadAbortException e) {
                throw new LoginException("Skype process has died unexpectedly", e);
            } finally {
                _canThrowSkypeExitException = false;
            }
        }

        private void StartSkypeProcess() {
            string arguments = string.Format("{0} /username:{1}", _arguments, _username);
            _skypeProcess = Process.Start(_skypeExePath, arguments);
        }

        private IntPtr GetInternetExplorerServerHandle(IntPtr loginWindowHandle) {
            // Get all child windows
            List<IntPtr> childWindows = new List<IntPtr>();
            WinApi.User32.EnumChildWindows(
                loginWindowHandle,
                (IntPtr wnd, ref IntPtr param) => {
                    childWindows.Add(wnd);
                    return true;
                },
                ref loginWindowHandle);

            // Check if any of child windows have Internet Explorer_Server class
            foreach (IntPtr childWindow in childWindows) {
                string className = GetWindowClassName(childWindow);
                if (className.IndexOf("Internet Explorer_Server", StringComparison.InvariantCulture) == 0)
                    return childWindow;
            }

            return IntPtr.Zero;
        }

        private GetSkypeLoginWindowDataResult GetSkypeLoginWindowData() {
            List<IntPtr> windowHandles = new List<IntPtr>();

            // Get all windows of the Skype process
            foreach (ProcessThread thread in _skypeProcess.Threads) {
                WinApi.User32.EnumThreadWindows(
                    (uint) thread.Id,
                    (hWnd, lParam) => {
                        windowHandles.Add(hWnd);
                        return true;
                    },
                    IntPtr.Zero);
            }

            // Find a window with class TLoginForm
            IntPtr loginWindowHandle = IntPtr.Zero;
            IntPtr skypeMainWindowHandle = IntPtr.Zero;
            foreach (IntPtr handle in windowHandles) {
                string windowClassName = GetWindowClassName(handle);
                switch (windowClassName) {
                    case "tSkMainForm":
                        skypeMainWindowHandle = handle;
                        continue;
                    case "TLoginForm":
                        loginWindowHandle = handle;
                        continue;
                }
            }

            if (skypeMainWindowHandle != IntPtr.Zero) {
                if (WinApi.User32.IsWindowVisible(skypeMainWindowHandle))
                    return new GetSkypeLoginWindowDataResult(GetSkypeLoginWindowDataResult.State.AlreadyLoggedIn);

                string contactListWindowTitle = GetWindowText(skypeMainWindowHandle);
                if (contactListWindowTitle != "Skype™‎" && contactListWindowTitle != "Skype")
                    return new GetSkypeLoginWindowDataResult(GetSkypeLoginWindowDataResult.State.AlreadyLoggedIn);
            }

            if (loginWindowHandle == IntPtr.Zero)
                return new GetSkypeLoginWindowDataResult(GetSkypeLoginWindowDataResult.State.Failed);

            // For some reason, Skype creates multiple windows with class TLoginForm,
            // and then destroys one of them. We only need the window that has Internet Explorer_Server,
            // so that's what we will search for.

            // Get all child windows
            List<IntPtr> childWindows = new List<IntPtr>();
            WinApi.User32.EnumChildWindows(
                loginWindowHandle,
                (IntPtr wnd, ref IntPtr param) => {
                    childWindows.Add(wnd);
                    return true;
                },
                ref loginWindowHandle);

            // Check if any of child windows has Internet Explorer_Server class
            foreach (IntPtr childWindow in childWindows) {
                string className = GetWindowClassName(childWindow);
                if (className.IndexOf("Edit", StringComparison.InvariantCulture) == 0) {
#if DEBUG
                    WinApi.User32.SetWindowText(loginWindowHandle, "SkypeLauncher: " + _username);
#endif
                    return
                        new GetSkypeLoginWindowDataResult(GetSkypeLoginWindowDataResult.State.Success) {
                            LoginWindowPtr = loginWindowHandle,
                            LoginEditPtr = childWindow
                        };
                }
            }

            return new GetSkypeLoginWindowDataResult(GetSkypeLoginWindowDataResult.State.Failed);
        }

        private void WebBrowserLogin() {
            Debug.WriteLine("Login URL\r\n\r\n{0}\r\n", (object) _loginBrowserHtmlDocument.location.href);

            // Wait for login.live.com page
            Debug.WriteLine("Waiting for host to switch to login.live.com");
            bool success = TimedOutOperation(kLoginPageReadyTimeout, kOperationRetryDelay, () => {
                try {
                    return
                        _loginBrowserHtmlDocument.location.host.EndsWith("login.live.com") &&
                        _loginBrowserHtmlDocument.location.pathname.StartsWith("/oauth20_authorize.srf");
                } catch (Exception) {
                    return false;
                }
            });

            if (!success)
                throw new LoginException("Failed waiting for login.live.com login page");

            // Wait for login page to load
            Debug.WriteLine("Wait until login.live.com login page is loaded");
            success = TimedOutOperation(kLoginPageReadyTimeout, kOperationRetryDelay, () => _loginBrowserHtmlDocument.readyState == "complete");
            if (!success)
                throw new LoginException("Unable to load the login.live.com login page");

            IHTMLElementCollection passwdElementCollection = null;
            success = TimedOutOperation(kLoginPageReadyTimeout, kOperationRetryDelay, () => {
                passwdElementCollection = _loginBrowserHtmlDocument.getElementsByName("passwd");
                return passwdElementCollection.length == 1;
            });

            if (!success)
                throw new LoginException("Unable to get password field element");

            // Get the login page elements
            Debug.WriteLine("Get the login page elements");
            HTMLInputElementClass passwordInputElement = (HTMLInputElementClass) ((IHTMLElementCollection3) passwdElementCollection).namedItem("passwd");
            HTMLInputElementClass signInButtonElement = (HTMLInputElementClass) _loginBrowserHtmlDocument.getElementById("idSIButton9");

            if (passwordInputElement == null || signInButtonElement == null)
                throw new LoginException("Unable to get login form elements");

            // Update password
            LoginBrowserSetInputFieldValue(passwordInputElement, _password);

            // Click Sign In button
            Debug.WriteLine("Click Sign In button");
            signInButtonElement.click();
        }

        private void LoginBrowserSetInputFieldValue(HTMLInputElementClass inputFieldElement, string value) {
            inputFieldElement.setAttribute("value", value);

            // Fire events
            RunOnStaThread(() => { LoginBrowserDispatchEvent("change", inputFieldElement.id); });
        }

        private void LoginBrowserDispatchEvent(string eventName, string elementId) {
            string script = @"
{
var __sk__evt = document.createEvent('HTMLEvents');
__sk__evt.initEvent('" + eventName + @"', false, true);
document.getElementById('" + elementId + @"').dispatchEvent(__sk__evt);
}
";

            _loginBrowserHtmlDocument.parentWindow.execScript(script);
        }

        private static void RunOnStaThread(Action action) {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA) {
                action();
                return;
            }

            AutoResetEvent resumeEvent = new AutoResetEvent(false);
            Exception threadException = null;
            Thread thread = new Thread(() => {
                try {
                    action();
                } catch (Exception e) {
                    threadException = e;
                } finally {
                    resumeEvent.Set();
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            resumeEvent.WaitOne();
            if (threadException != null)
                throw new LoginException("Unexpected error", threadException);
        }

        private static bool TimedOutOperation(int timeout, int retryDelay, Func<bool> operation) {
            int retries = timeout / retryDelay;
            for (int i = 0; i < retries; i++) {
                if (operation())
                    return true;

                Thread.Sleep(retryDelay);
            }

            return false;
        }

        private static string GetWindowText(IntPtr windowHandle) {
            int textLen = WinApi.User32.GetWindowTextLength(windowHandle);
            StringBuilder sb = new StringBuilder(textLen + 1);
            WinApi.User32.GetWindowText(windowHandle, sb, sb.Capacity);
            string windowTitle = sb.ToString();
            return windowTitle;
        }

        private static string GetWindowClassName(IntPtr hWnd) {
            StringBuilder classNameBuffer = new StringBuilder(128);
            WinApi.User32.GetClassName(hWnd, classNameBuffer, classNameBuffer.Capacity);
            string className = classNameBuffer.ToString();
            return className;
        }

        private static HTMLDocumentClass GetHtmlDocumentClassFromInternetExplorerServerHandle(IntPtr internetExplorerServerHandle) {
            int lngMsg = WinApi.User32.RegisterWindowMessage("WM_HTML_GETOBJECT");
            if (lngMsg == 0)
                return null;

            int lRes;
            WinApi.User32.SendMessageTimeout(internetExplorerServerHandle, lngMsg, 0, 0, NativeMethods.SMTO_ABORTIFHUNG, kHtmlGetObjectTimeout, out lRes);
            if (lRes == 0)
                return null;

            HTMLDocumentClass document = null;
            Guid guid = typeof(IHTMLDocument2).GUID;
            int hResult = NativeMethods.ObjectFromLresult(lRes, ref guid, 0, ref document);
            if (hResult != 0)
                return null;

            return document;
        }

        private struct GetSkypeLoginWindowDataResult {
            public readonly State LoginState;
            public IntPtr LoginWindowPtr;
            public IntPtr LoginEditPtr;

            public GetSkypeLoginWindowDataResult(State loginState) : this() {
                LoginState = loginState;
            }

            public enum State {
                Success,
                Failed,
                AlreadyLoggedIn
            }
        }

        private static class NativeMethods {
            public const int SMTO_ABORTIFHUNG = 0x2;

            [DllImport("OLEACC.dll")]
            public static extern int ObjectFromLresult(int lResult, ref Guid riid, int wParam, ref HTMLDocumentClass ppvObject);
        }
    }
}
