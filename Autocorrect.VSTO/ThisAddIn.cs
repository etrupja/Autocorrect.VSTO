using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;

namespace Autocorrect.VSTO
{
    public partial class ThisAddIn
    {
        //private SafeNativeMethods.HookProc _mouseProc;
        private SafeNativeMethods.HookProc _keyboardProc;

        //private IntPtr _hookIdMouse;
        private IntPtr _hookIdKeyboard;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //_mouseProc = MouseHookCallback;
            _keyboardProc = KeyboardHookCallback;

            SetWindowsHooks();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            UnhookWindowsHooks();
        }

        private void SetWindowsHooks()
        {
            uint threadId = (uint)SafeNativeMethods.GetCurrentThreadId();

            //_hookIdMouse =
            //    SafeNativeMethods.SetWindowsHookEx(
            //        (int)SafeNativeMethods.HookType.WH_MOUSE,
            //        _mouseProc,
            //        IntPtr.Zero,
            //        threadId);

            _hookIdKeyboard =
                SafeNativeMethods.SetWindowsHookEx(
                    (int)SafeNativeMethods.HookType.WH_KEYBOARD,
                    _keyboardProc,
                    IntPtr.Zero,
                    threadId);
        }

        private void UnhookWindowsHooks()
        {
            SafeNativeMethods.UnhookWindowsHookEx(_hookIdKeyboard);
            //SafeNativeMethods.UnhookWindowsHookEx(_hookIdMouse);
        }

        //private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        //{
        //    if (nCode >= 0)
        //    {
        //        var mouseHookStruct =
        //            (SafeNativeMethods.MouseHookStructEx)
        //                Marshal.PtrToStructure(lParam, typeof(SafeNativeMethods.MouseHookStructEx));

        //        // handle mouse message here
        //        var message = (SafeNativeMethods.WindowMessages)wParam;
        //        Debug.WriteLine(
        //            "{0} event detected at position {1} - {2}",
        //            message,
        //            mouseHookStruct.pt.X,
        //            mouseHookStruct.pt.Y);
        //        //MessageBox.Show(
        //        //    $"{message} event detected at position {mouseHookStruct.pt.X} - {mouseHookStruct.pt.Y}");
        //    }
        //    return SafeNativeMethods.CallNextHookEx(
        //        _hookIdKeyboard,
        //        nCode,
        //        wParam,
        //        lParam);
        //}

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //Keys key = Keys.Space;
            //if ((uint)wParam == (uint)key)
            //{
            //    MessageBox.Show("You pressed spacebar - wParam");
            //}

            //if (nCode >= 0)
            //{
            //    Debug.WriteLine("Key event detected.");
            //}

            // Feel free to move the const to a private field.
            const int HC_ACTION = 0;
            if (nCode == HC_ACTION)
            {
                Keys key = (Keys)wParam;
                KeyEventArgs args = new KeyEventArgs(key);

                bool isKeyDown = ((ulong)lParam & 0x40000000) == 0;
                bool isSpaceKey = args.KeyCode == Keys.Space;
                if (isKeyDown && isSpaceKey)
                    onKeyDown(args);
                //else
                //{
                //    bool isLastKeyUp = ((ulong)lParam & 0x80000000) == 0x80000000;
                //    if (isLastKeyUp)
                //        onKeyUp(args);
                //}
            }

            return SafeNativeMethods.CallNextHookEx(
                _hookIdKeyboard,
                nCode,
                wParam,
                lParam);
        }

        private void onKeyUp(KeyEventArgs args)
        {
            MessageBox.Show($"onKeyUp - {args.KeyCode.ToString()}. Text - {Globals.ThisAddIn.Application.Selection.Text}");
        }

        private void onKeyDown(KeyEventArgs args)
        {
            Word._Document oDoc = Globals.ThisAddIn.Application.ActiveDocument;
            int start = oDoc.Content.Start;
            int end = oDoc.Content.End;

            string text = oDoc.Range(start, end).Text;
            string lastWord = text.Split(' ').Last();
            var finalResult = lastWord.Replace("\r", "");

            switch (finalResult)
            {
                case "eshte":
                    oDoc.Range(text.IndexOf(" ") + 1, end).Text = "është";
                    break;
                case "nje":
                    oDoc.Range(text.IndexOf(" ") + 1, end).Text = "një";
                    break;
                case "per":
                    oDoc.Range(text.IndexOf(" ") + 1, end).Text = "për";
                    break;
                case "te":
                    oDoc.Range(text.IndexOf(" ") + 1, end).Text = "të";
                    break;
                case "pare":
                    oDoc.Range(text.IndexOf(" ") + 1, end).Text = "parë";
                    break;
                case "funksionoje":
                    oDoc.Range(text.IndexOf(" ") + 1, end).Text = "funksionojë";
                    break;


            }

            //end = oDoc.Content.End;

            //var rng = oDoc.Range(start, end);
            //object NewEndPos = oDoc.Range(text.LastIndexOf(" ") + 1, end).StoryLength - 1;
            //int lastIndex = text.LastIndexOf("\r");
            //oDoc.Range(lastIndex - 1, lastIndex).Select();
            //rng.Select();44
            //oDoc.Range(start, end).Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            //oDoc.Range(end-1, end).Select();

        }




        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }

    internal static class SafeNativeMethods
    {
        public delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        public enum HookType
        {
            WH_KEYBOARD = 2,
            WH_MOUSE = 7
        }

        public enum WindowMessages : uint
        {
            WM_KEYDOWN = 0x0100,
            WM_KEYFIRST = 0x0100,
            WM_KEYLAST = 0x0108,
            WM_KEYUP = 0x0101,
            WM_LBUTTONDBLCLK = 0x0203,
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_MBUTTONDBLCLK = 0x0209,
            WM_MBUTTONDOWN = 0x0207,
            WM_MBUTTONUP = 0x0208,
            WM_MOUSEACTIVATE = 0x0021,
            WM_MOUSEFIRST = 0x0200,
            WM_MOUSEHOVER = 0x02A1,
            WM_MOUSELAST = 0x020D,
            WM_MOUSELEAVE = 0x02A3,
            WM_MOUSEMOVE = 0x0200,
            WM_MOUSEWHEEL = 0x020A,
            WM_MOUSEHWHEEL = 0x020E,
            WM_RBUTTONDBLCLK = 0x0206,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205,
            WM_SYSDEADCHAR = 0x0107,
            WM_SYSKEYDOWN = 0x0104,
            WM_SYSKEYUP = 0x0105
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(
            int idHook,
            HookProc lpfn,
            IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr CallNextHookEx(
            IntPtr hhk,
            int nCode,
            IntPtr wParam,
            IntPtr lParam);

        [DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        public struct Point
        {
            public int X;
            public int Y;

            public Point(int x, int y)
            {
                X = x;
                Y = y;
            }

            public static implicit operator System.Drawing.Point(Point p)
            {
                return new System.Drawing.Point(p.X, p.Y);
            }

            public static implicit operator Point(System.Drawing.Point p)
            {
                return new Point(p.X, p.Y);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MouseHookStructEx
        {
            public Point pt;
            public IntPtr hwnd;
            public uint wHitTestCode;
            public IntPtr dwExtraInfo;
            public int MouseData;
        }
    }
}
