using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace IEPop
{
    public static class FlashWindow
    {
        private static readonly Timer disableFlashing;
        private static readonly List<System.Windows.Forms.Form> flashingForms;

        static FlashWindow()
        {
            flashingForms = new List<System.Windows.Forms.Form>();
            disableFlashing = new Timer(StopAllFlash);
        }

        private static void StopAllFlash(object p1)
        {
            try
            {
                lock (flashingForms)
                {
                    var forms = new List<System.Windows.Forms.Form>(flashingForms);
                    flashingForms.Clear();
                    foreach (var f in forms)
                    {
                        Stop(f);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message);  
            }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)]
        internal struct FLASHWINFO
        {
            /// <summary>
            /// The size of the structure in bytes.
            /// </summary>
            public uint cbSize;
            /// <summary>
            /// A Handle to the Window to be Flashed. The window can be either opened or minimized.
            /// </summary>
            public IntPtr hwnd;
            /// <summary>
            /// The Flash Status.
            /// </summary>
            public uint dwFlags;
            /// <summary>
            /// The number of times to Flash the window.
            /// </summary>
            public uint uCount;
            /// <summary>
            /// The rate at which the Window is to be flashed, in milliseconds. If Zero, the function uses the default cursor blink rate.
            /// </summary>
            public uint dwTimeout;
        }

        /// <summary>
        /// Stop flashing. The system restores the window to its original stae.
        /// </summary>
        public const uint FLASHW_STOP = 0;

        /// <summary>
        /// Flash the window caption.
        /// </summary>
        public const uint FLASHW_CAPTION = 1;

        /// <summary>
        /// Flash the taskbar button.
        /// </summary>
        public const uint FLASHW_TRAY = 2;

        /// <summary>
        /// Flash both the window caption and taskbar button.
        /// This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags.
        /// </summary>
        public const uint FLASHW_ALL = 3;

        /// <summary>
        /// Flash continuously, until the FLASHW_STOP flag is set.
        /// </summary>
        public const uint FLASHW_TIMER = 4;

        /// <summary>
        /// Flash continuously until the window comes to the foreground.
        /// </summary>
        public const uint FLASHW_TIMERNOFG = 12;


        /// <summary>
        /// Flash the spacified Window (Form) until it recieves focus.
        /// </summary>
        /// <param name="form">The Form (Window) to Flash.</param>
        /// <returns></returns>
        public static bool Flash(System.Windows.Forms.Form form)
        {
            disableFlashing.Change(60000, Timeout.Infinite);
            lock ( flashingForms )
            {
                if ( !flashingForms.Contains(form) )
                {
                    flashingForms.Add(form);
                }
            }
            if ( form.InvokeRequired )
            {
                return (bool)form.Invoke(new InvokeFormUpdateDelegate(f =>
                {
                    FLASHWINFO fi2 = Create_FLASHWINFO(form.Handle, FLASHW_ALL | FLASHW_TIMERNOFG, uint.MaxValue, 0);
                    return FlashWindowEx(ref fi2);
                }), form);
            }            
            FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL | FLASHW_TIMERNOFG, uint.MaxValue, 0);
            return FlashWindowEx(ref fi);
        }

        internal static FLASHWINFO Create_FLASHWINFO(IntPtr handle, uint flags, uint count, uint timeout)
        {
            FLASHWINFO fi = new FLASHWINFO();
            fi.cbSize = Convert.ToUInt32(Marshal.SizeOf(fi));
            fi.hwnd = handle;
            fi.dwFlags = flags;
            fi.uCount = count;
            fi.dwTimeout = timeout;
            return fi;
        }

        /// <summary>
        /// Flash the specified Window (form) for the specified number of times
        /// </summary>
        /// <param name="form">The Form (Window) to Flash.</param>
        /// <param name="count">The number of times to Flash.</param>
        /// <returns></returns>
        public static bool Flash(System.Windows.Forms.Form form, uint count)
        {
            FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL, count, 0);
            return FlashWindowEx(ref fi);
        }

        /// <summary>
        /// Start Flashing the specified Window (form)
        /// </summary>
        /// <param name="form">The Form (Window) to Flash.</param>
        /// <returns></returns>
        public static bool Start(System.Windows.Forms.Form form)
        {
            disableFlashing.Change(60000, Timeout.Infinite);
            lock ( flashingForms )
            {
                if ( !flashingForms.Contains(form) )
                {
                    flashingForms.Add(form);
                }
            }
            if (form.InvokeRequired)
            {
                return (bool)form.Invoke(new InvokeFormUpdateDelegate(f =>
                {
                    FLASHWINFO fi2 = Create_FLASHWINFO(form.Handle, FLASHW_ALL, uint.MaxValue, 0);
                    return FlashWindowEx(ref fi2);
                }), form);
            }
            FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_ALL, uint.MaxValue, 0);
            return FlashWindowEx(ref fi);
        }

        /// <summary>
        /// Stop Flashing the specified Window (form)
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        public static bool Stop(System.Windows.Forms.Form form)
        {
            lock ( flashingForms )
            {
                if ( flashingForms.Contains(form) )
                {
                    flashingForms.Remove(form);
                }
            }
            if (form.InvokeRequired)
            {
                return (bool)form.Invoke(new InvokeFormUpdateDelegate(f =>
                {
                    FLASHWINFO fi2 = Create_FLASHWINFO(form.Handle, FLASHW_STOP, uint.MaxValue, 0);
                    return FlashWindowEx(ref fi2);
                }), form);
            }

            FLASHWINFO fi = Create_FLASHWINFO(form.Handle, FLASHW_STOP, uint.MaxValue, 0);
            return FlashWindowEx(ref fi);
        }

        private delegate bool InvokeFormUpdateDelegate(System.Windows.Forms.Form form);
    }
}
