using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace Autocorrect.Common
{
    public  class AddinHelper
    {
        public EventHandler<KeyEventArgs> OnKeyUp;
        private SafeNativeMethods.HookProc _keyboardProc;
        private IntPtr _hookIdKeyboard;
        public  void RegisterEvents()
        {
            _keyboardProc = KeyboardHookCallback;
            SetWindowsHooks();
        }
        public void UnRegisterEvents()
        {
            UnhookWindowsHooks();
        }
        private  void SetWindowsHooks()
        {

            uint threadId = (uint)SafeNativeMethods.GetCurrentThreadId();
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
        }
        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            const int HC_ACTION = 0;
            if (nCode == HC_ACTION)
            {
                Keys key = (Keys)wParam;
                KeyEventArgs args = new KeyEventArgs(key);
                bool isKeyDown = ((ulong)lParam & 0x40000000) == 0;
                //bool isSpaceKey = args.KeyCode == Keys.Space;
                if (!isKeyDown) OnKeyDownHandler(args);
            }

            return SafeNativeMethods.CallNextHookEx(
                _hookIdKeyboard,
                nCode,
                wParam,
                lParam);
        }

        private void OnKeyDownHandler(KeyEventArgs args)
        {
            this.OnKeyUp?.Invoke(this,args);
        }
    }
}
