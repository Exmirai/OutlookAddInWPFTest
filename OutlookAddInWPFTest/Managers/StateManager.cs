using System;
using System.ComponentModel;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Utils;

using System.Runtime.InteropServices;
using OutlookAddInWPFTest.Forms;
using OutlookAddInWPFTest.Forms.JudicoWindow;

namespace OutlookAddInWPFTest.Managers
{
    public static class StateManager
    {
        private static WinAPI.HookProc _cbtProc = CBTHook;
        private static WinAPI.WinEventDelegate _winEventProc = winEvProc;
        private static IntPtr _cbtHook;
        private static IntPtr _winEventHook;
        public static UIStateEnum UiState { get; set; }
        public static OutlookStateEnum OutlookState { get; set; }

        public static void Init()
        {
          //  if ((_cbtHook = WinAPI.SetWindowsHookEx(WinAPI.HookType.WH_CBT, _cbtProc, IntPtr.Zero, WinAPI.GetCurrentThreadId())) == IntPtr.Zero)
         //   {
          //      throw new Win32Exception(Marshal.GetLastWin32Error());
           // }
            if ((_winEventHook = WinAPI.SetWinEventHook(WinAPI.WinEvents.EVENT_SYSTEM_FOREGROUND, WinAPI.WinEvents.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, _winEventProc, 0, 0, WinAPI.WinEventFlags.WINEVENT_OUTOFCONTEXT )) == IntPtr.Zero)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            UiState = UIStateEnum.NONE;
            OutlookState = OutlookStateEnum.MINIMIZED;
        }

        public static IntPtr CBTHook(int nCode, IntPtr wParam, IntPtr lParam)
        {
            var outlookHwnd = OutlookUtils.GetOutlookWindow();
            var wordHwnd = OutlookUtils.GetWordWindow();
            if (JButton.Instance == null)
            {
                return WinAPI.CallNextHookEx(_cbtHook, nCode, wParam, lParam);
            }
            var jButtonHwnd = new System.Windows.Interop.WindowInteropHelper(JButton.Instance).Handle;
            switch ((WinAPI.HCBT)nCode)
            {
                case WinAPI.HCBT.Activate:
                    if (wParam == outlookHwnd || wParam == wordHwnd || wParam == jButtonHwnd ||WinAPI.GetWindow(wParam, WinAPI.GetWindowType.GW_OWNER) == outlookHwnd)
                    {
                        var payload =
                            (WinAPI.CBTACTIVATESTRUCT)Marshal.PtrToStructure(lParam, typeof(WinAPI.CBTACTIVATESTRUCT));
                        OutlookState = OutlookStateEnum.INBOX;
                    }
                    else
                    {
                        OutlookState = OutlookStateEnum.MINIMIZED;
                    }

                    break;
                case WinAPI.HCBT.MinMax:
                    var showWndCmd = (WinAPI.ShowWindowCommands)lParam;
                    if (wParam == outlookHwnd || wParam == wordHwnd || wParam == jButtonHwnd || WinAPI.GetWindow(wParam, WinAPI.GetWindowType.GW_OWNER) == outlookHwnd)
                    {
                       // OutlookState = OutlookStateEnum.INBOX;
                    }
                    else
                    {
                        //OutlookState = OutlookStateEnum.MINIMIZED;
                    }
                    break;
                case WinAPI.HCBT.MoveSize:
                    break;
                case WinAPI.HCBT.SetFocus:
                    if (wParam == wordHwnd)
                    {
                        //OutlookState = OutlookStateEnum.INBOX;
                    }
                    break;
                default:
                    break;
            }

            return WinAPI.CallNextHookEx(_cbtHook, nCode, wParam, lParam);
        }

        public static void winEvProc(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (eventType != (uint)WinAPI.WinEvents.EVENT_SYSTEM_FOREGROUND)
            {
                return;
            }
            var outlookHwnd = OutlookUtils.GetOutlookWindow();
            var wordHwnd = OutlookUtils.GetWordWindow();
            if (JButton.Instance == null || Overlay.Instance == null || JudicoWindow.Instance == null)
            {
                return;
            }
            var jButtonHwnd = new System.Windows.Interop.WindowInteropHelper(JButton.Instance).Handle;
            var overlayHwnd = new System.Windows.Interop.WindowInteropHelper(Overlay.Instance).Handle;
            var jWindowHwnd = new System.Windows.Interop.WindowInteropHelper(JudicoWindow.Instance).Handle;
            if (hwnd == outlookHwnd || 
                hwnd == wordHwnd || 
                hwnd == jButtonHwnd || 
                hwnd == jWindowHwnd ||
                hwnd == overlayHwnd /* || WinAPI.GetWindow(hwnd, WinAPI.GetWindowType.GW_OWNER) == outlookHwnd || WinAPI.GetWindow(hwnd, WinAPI.GetWindowType.GW_OWNER) == wordHwnd*/)
            {
                OutlookState = OutlookStateEnum.INBOX;
            }
            else
            {
                OutlookState = OutlookStateEnum.MINIMIZED;
            }
        }
    }
}
