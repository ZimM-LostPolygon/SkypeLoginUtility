using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using mshtml;

namespace SkypeConsoleLogin {
    internal class SkypeLauncher {
        private const int kOperationRetryDelay = 30;
        private const int kSkypeRestartDelay = 1000;
        private const int kSkypeMaxRestarts = 3;
        private const int kWaitForSkypeLoginWindowTimeout = 12000;
        private const int kHtmlGetObjectTimeout = 1000;
        private const int kLoginPageReadyTimeout = 10000;
        private readonly string _arguments;
        private readonly string _password;

        private readonly string _skypeExePath;
        private readonly string _username;
        private HTMLDocumentClass _browserHtmlDocument;
        private Process _skypeProcess;
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
        }

        public void Launch() {
            _currentThread = Thread.CurrentThread;
            StartSkypeProcess();

            Thread skypeWaitForExitThread = new Thread(() => {
                _skypeProcess.WaitForExit();
                if (_canThrowSkypeExitException) {
                    _currentThread.Abort();
                }
            });
            skypeWaitForExitThread.IsBackground = true;
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

            IntPtr internetExplorerServerHandle = IntPtr.Zero;
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

                    IntPtr internetExplorerServerHandleTemp;
                    GetSkypeLoginWindowDataResult result = GetSkypeLoginWindowData(out internetExplorerServerHandleTemp);

                    switch (result) {
                        case GetSkypeLoginWindowDataResult.Success:
                            internetExplorerServerHandle = internetExplorerServerHandleTemp;
                            return true;
                        case GetSkypeLoginWindowDataResult.Failed:
                            return false;
                        case GetSkypeLoginWindowDataResult.AlreadyLoggedIn:
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

            _canThrowSkypeExitException = true;
            try {
                // Get HTMLDocumentClass of the login page
                Debug.WriteLine("Get HTMLDocumentClass of the login page");
                _browserHtmlDocument = null;
                success = TimedOutOperation(kWaitForSkypeLoginWindowTimeout, kOperationRetryDelay, () => {
                    _browserHtmlDocument = GetHtmlDocumentClassFromInternetExplorerServerHandle(internetExplorerServerHandle);
                    return _browserHtmlDocument != null;
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

        private void WebBrowserLogin() {
            Debug.WriteLine("Login URL\r\n\r\n{0}\r\n", (object) _browserHtmlDocument.location.href);

            // Wait until login page is loaded
            Debug.WriteLine("Wait until skype.com login page is loaded");
            bool success = TimedOutOperation(kLoginPageReadyTimeout, kOperationRetryDelay, () => _browserHtmlDocument.readyState == "complete");
            if (!success)
                throw new LoginException("Unable to load the skype.com login page");

            // Get the login page elements
            Debug.WriteLine("Get the login page elements");
            HTMLInputElementClass usernameInputFieldElement = (HTMLInputElementClass) _browserHtmlDocument.getElementById("unifiedUsername");
            IHTMLElement signInButtonElement = _browserHtmlDocument.getElementById("unifiedSignIn");

            if (usernameInputFieldElement == null || signInButtonElement == null)
                throw new LoginException("Unable to get login form elements");

            // Update username
            SetLoginInputFieldValue(usernameInputFieldElement, _username);

            // Click Sign In button
            Debug.WriteLine("Click Sign In button");
            signInButtonElement.click();

            // Wait for login.live.com page
            Debug.WriteLine("Waiting for host to switch to login.live.com");
            success = TimedOutOperation(kLoginPageReadyTimeout, kOperationRetryDelay, () => {
                try {
                    return
                        _browserHtmlDocument.location.host.EndsWith("login.live.com") &&
                        _browserHtmlDocument.location.pathname.StartsWith("/oauth20_authorize.srf");
                } catch (Exception) {
                    return false;
                }
            });

            if (!success)
                throw new LoginException("Failed waiting for login.live.com login page");

            // Wait for login page to load
            Debug.WriteLine("Wait until login.live.com login page is loaded");
            success = TimedOutOperation(kLoginPageReadyTimeout, kOperationRetryDelay, () => _browserHtmlDocument.readyState == "complete");
            if (!success)
                throw new LoginException("Unable to load the login.live.com login page");

            Debug.WriteLine("Login URL\r\n\r\n{0}\r\n", (object) _browserHtmlDocument.location.href);

            IHTMLElementCollection passwdElementCollection = null;
            success = TimedOutOperation(kLoginPageReadyTimeout, kOperationRetryDelay, () => {
                passwdElementCollection = _browserHtmlDocument.getElementsByName("passwd");
                return passwdElementCollection.length == 1;
            });

            if (!success)
                throw new LoginException("Unable to get password field element");

            // Get the login page elements
            Debug.WriteLine("Get the login page elements");
            HTMLInputElementClass passwordInputElement = (HTMLInputElementClass) ((IHTMLElementCollection3) passwdElementCollection).namedItem("passwd");
            signInButtonElement = _browserHtmlDocument.getElementById("idSIButton9");

            if (passwordInputElement == null || signInButtonElement == null)
                throw new LoginException("Unable to get login form elements");

            // Update password
            SetLoginInputFieldValue(passwordInputElement, _password);

            // Click Sign In button
            Debug.WriteLine("Click Sign In button");
            signInButtonElement.click();
        }

        private void StartSkypeProcess() {
            string arguments = string.Format("{0} /username:{1}", _arguments, _username);
            _skypeProcess = Process.Start(_skypeExePath, arguments);
        }

        private GetSkypeLoginWindowDataResult GetSkypeLoginWindowData(out IntPtr internetExplorerServerHandle) {
            internetExplorerServerHandle = IntPtr.Zero;
            List<IntPtr> windowHandles = new List<IntPtr>();

            // Get all windows of the Skype process
            foreach (ProcessThread thread in _skypeProcess.Threads) {
                NativeMethods.EnumThreadWindows(
                    thread.Id,
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
                if (NativeMethods.IsWindowVisible(skypeMainWindowHandle))
                    return GetSkypeLoginWindowDataResult.AlreadyLoggedIn;

                string contactListWindowTitle = GetWindowText(skypeMainWindowHandle);
                if (contactListWindowTitle != "Skype™‎" && contactListWindowTitle != "Skype")
                    return GetSkypeLoginWindowDataResult.AlreadyLoggedIn;
            }

            if (loginWindowHandle == IntPtr.Zero)
                return GetSkypeLoginWindowDataResult.Failed;

            // For some reason, Skype creates multiple windows with class TLoginForm,
            // and then destroys one of them. We only need the window that has Internet Explorer_Server,
            // so that's what we will search for.

            // Get all child windows
            List<IntPtr> childWindows = new List<IntPtr>();
            NativeMethods.EnumChildWindows(
                loginWindowHandle,
                (IntPtr wnd, ref IntPtr param) => {
                    childWindows.Add(wnd);
                    return 1;
                },
                ref loginWindowHandle);

            // Check if any of child windows has Internet Explorer_Server class
            foreach (IntPtr childWindow in childWindows) {
                string className = GetWindowClassName(childWindow);
                if (className.IndexOf("Internet Explorer_Server", StringComparison.InvariantCulture) == 0) {
                    internetExplorerServerHandle = childWindow;
#if DEBUG
                    NativeMethods.SetWindowText(loginWindowHandle, "SkypeLauncher: " + _username);
#endif
                    return GetSkypeLoginWindowDataResult.Success;
                }
            }

            return GetSkypeLoginWindowDataResult.Failed;
        }

        private void SetLoginInputFieldValue(HTMLInputElementClass inputFieldElement, string value) {
            inputFieldElement.setAttribute("value", value);

            // Fire events
            RunOnStaThread(() => { DispatchEvent("change", inputFieldElement.id); });
        }

        private void DispatchEvent(string eventName, string elementId) {
            string script = @"
{
var __sk__evt = document.createEvent('HTMLEvents');
__sk__evt.initEvent('" + eventName + @"', false, true);
document.getElementById('" + elementId + @"').dispatchEvent(__sk__evt);
}
";

            _browserHtmlDocument.parentWindow.execScript(script);
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

        private static HTMLDocumentClass GetHtmlDocumentClassFromInternetExplorerServerHandle(IntPtr internetExplorerServerHandle) {
            int lngMsg = NativeMethods.RegisterWindowMessage("WM_HTML_GETOBJECT");
            if (lngMsg == 0)
                return null;

            int lRes;
            NativeMethods.SendMessageTimeout(internetExplorerServerHandle, lngMsg, 0, 0, NativeMethods.SMTO_ABORTIFHUNG, kHtmlGetObjectTimeout, out lRes);
            if (lRes == 0)
                return null;

            HTMLDocumentClass document = null;
            Guid guid = typeof(IHTMLDocument2).GUID;
            int hResult = NativeMethods.ObjectFromLresult(lRes, ref guid, 0, ref document);
            if (hResult != 0)
                return null;

            return document;
        }

        private static string GetWindowText(IntPtr windowHandle) {
            int textLen = NativeMethods.GetWindowTextLength(windowHandle);
            StringBuilder sb = new StringBuilder(textLen + 1);
            NativeMethods.GetWindowText(windowHandle, sb, sb.Capacity);
            string windowTitle = sb.ToString();
            return windowTitle;
        }

        private static string GetWindowClassName(IntPtr hWnd) {
            StringBuilder classNameBuffer = new StringBuilder(128);
            NativeMethods.GetClassName(hWnd, classNameBuffer, classNameBuffer.Capacity);
            string className = classNameBuffer.ToString();
            return className;
        }

        private enum GetSkypeLoginWindowDataResult {
            Success,
            Failed,
            AlreadyLoggedIn
        }

        private static class NativeMethods {
            public const int SMTO_ABORTIFHUNG = 0x2;

            public delegate int EnumProc(IntPtr hWnd, ref IntPtr lParam);

            public delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

            [DllImport("user32.dll")]
            public static extern bool IsWindowVisible(IntPtr hwnd);

            [DllImport("user32.dll", EntryPoint = "GetClassNameA")]
            public static extern int GetClassName(IntPtr hwnd, StringBuilder lpClassName, int nMaxCount);

            [DllImport("user32.dll")]
            public static extern int EnumChildWindows(IntPtr hWndParent, EnumProc lpEnumFunc, ref IntPtr lParam);

            [DllImport("user32.dll")]
            public static extern int EnumChildWindows(IntPtr hWndParent, EnumProc lpEnumFunc, IntPtr lParam);

            [DllImport("user32.dll")]
            public static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

            [DllImport("user32.dll", EntryPoint = "RegisterWindowMessageA")]
            public static extern int RegisterWindowMessage(string lpString);

            [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutA")]
            public static extern int SendMessageTimeout(IntPtr hwnd, int msg, int wParam, int lParam, int fuFlags, int uTimeout, out int lpdwResult);

            [DllImport("user32.dll", CharSet=CharSet.Unicode)]
            public static extern int GetWindowTextLength(IntPtr hWnd);

            [DllImport("user32.dll", CharSet=CharSet.Unicode)]
            public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxLength);

            [DllImport("user32.dll", CharSet=CharSet.Unicode)]
            public static extern int SetWindowText(IntPtr hWnd, string text);

            [DllImport("OLEACC.dll")]
            public static extern int ObjectFromLresult(int lResult, ref Guid riid, int wParam, ref HTMLDocumentClass ppvObject);

            [DllImport("kernel32.dll")]
            public static extern uint GetCurrentThreadId();
        }
    }
}