using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SHDocVw;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace IEPop
{
    public partial class Form1 : Form
    {
        internal const string WindowName = "guidewire policycenter";
        public Form1()
        {
            InitializeComponent();
            WindowsInstance = new ShellWindows();

            Load += (s, e) => Close();
        }
        public Form1(params string[] args) : this()
        {
            string cleanargs;
            if (args.Length == 1)
            {
                cleanargs = args[0].Substring(6);
            }
            else { return; }
            //foreach (var arg in cleanargs.Split)
            {
                var arg = cleanargs;
                MessageBox.Show(arg);
                var kvp = arg.Split('=');
                if (kvp.Length == 2)
                {
                    if (kvp[0] == "screenpop")
                    {
                        ScreenPop(kvp[1]);
                    }
                    else if (kvp[0] == "flash")
                    {
                        StartFlash(kvp[1]);
                    }
                }
            }
        }
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        private void StartFlash(string window)
        {
            try
            {
                InternetExplorer targetWindow = null;
                for (int i = 0; i < WindowsInstance.Count; i++)
                {
                    var tempIe = WindowsInstance.Item(i) as InternetExplorer;

                    if (!IsIeAndNamedWindow(tempIe, window)) continue;

                    targetWindow = tempIe;
                    break;
                }
                if (targetWindow != null)
                {
                    uint pid;
                    GetWindowThreadProcessId((IntPtr)targetWindow.HWND, out pid);
                    Process ie = Process.GetProcessById((int)pid);
                    var flash = FlashWindow.Create_FLASHWINFO((IntPtr)targetWindow.HWND, FlashWindow.FLASHW_ALL | FlashWindow.FLASHW_TIMERNOFG, UInt16.MaxValue / 2, 0);
                    //var flash = FlashWindow.Create_FLASHWINFO(ie.MainWindowHandle, FlashWindow.FLASHW_ALL | FlashWindow.FLASHW_TIMERNOFG, UInt16.MaxValue / 2, 0);
                    var hr = FlashWindow.FlashWindowEx(ref flash);
                    var error = Marshal.GetLastWin32Error();
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void ScreenPop(string Url)
        {
            var url = DecodeUrlString(Url);
            if (!PushUrlToPolicyCenterWindow(url))
            {
                var wb = new System.Windows.Forms.WebBrowser();

                wb.Navigate(url, true);
                wb.Focus();
            }
        }
        private static string DecodeUrlString(string url)
        {
            string newUrl;
            while ((newUrl = Uri.UnescapeDataString(url)) != url)
                url = newUrl;
            return newUrl;
        }

        protected ShellWindows WindowsInstance;

        protected bool IsIeAndNamedWindow(InternetExplorer window, string windowName)
        {
            windowName = windowName.ToLowerInvariant();
            Regex ieReg = new Regex(@"Internet\b\s+Explorer\b", RegexOptions.IgnoreCase);

            if (window == null || !ieReg.IsMatch(window.Name)) return false;

            var doc = window.Document as mshtml.IHTMLDocument2;

            return doc != null && doc.title.ToLowerInvariant().Contains(windowName);
        }
        private bool PushUrlToPolicyCenterWindow(string url)
        {
            InternetExplorer targetWindow = null;
            bool pushSuccessful = false;

            try
            {
                for (int i = 0; i < WindowsInstance.Count; i++)
                {
                    var tempIe = WindowsInstance.Item(i) as InternetExplorer;

                    if (!IsIeAndNamedWindow(tempIe, WindowName)) continue;

                    targetWindow = tempIe;
                    break;
                }

                if (targetWindow != null)
                {
                    targetWindow.Navigate2(url, null, "_self", null, null);

                    try
                    {
                        var hWnd = (IntPtr)targetWindow.HWND;
                        TabActivator.ShowWindowAsync(hWnd, TabActivator.SW_SHOWMAXIMIZED);
                        TabActivator.SetForegroundWindow(hWnd);

                        new TabActivator(hWnd).ActivateByTabsUrl(url);
                    }
                    catch (Exception aex)
                    {
                        _logger.Warn("*** An error occurred. See details below. ***");
                        _logger.Error("*** {0} ***", aex.Message);
                        _logger.Error("*** {0} ***", aex.StackTrace);
                    }

                    pushSuccessful = true;
                }
            }
            catch (Exception ex)
            {
                pushSuccessful = false;

                _logger.Warn("*** An error occurred. See details below. ***");
                _logger.Error("*** {0} ***", ex.Message);
                _logger.Error("*** {0} ***", ex.StackTrace);
            }

            return pushSuccessful;
        }
    }
}

