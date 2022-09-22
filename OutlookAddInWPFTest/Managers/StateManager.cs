using System;
using System.ComponentModel;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Utils;

using System.Runtime.InteropServices;
using OutlookAddInWPFTest.Forms;

namespace OutlookAddInWPFTest.Managers
{
    public static class StateManager
    {
        private static WinAPI.HookProc _cbtProc = CBTHook;
        private static IntPtr _cbtHook;
        public static UIStateEnum UiState { get; set; }
        public static OutlookStateEnum OutlookState { get; set; }

        public static void Init()
        {
            if ((_cbtHook = WinAPI.SetWindowsHookEx(WinAPI.HookType.WH_CBT, _cbtProc, IntPtr.Zero, WinAPI.GetCurrentThreadId())) == IntPtr.Zero)
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
    }
}
