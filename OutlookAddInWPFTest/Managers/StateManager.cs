﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OutlookAddInWPFTest.Enum;
using OutlookAddInWPFTest.Utils;

using System.Runtime.InteropServices;
using System.Threading;

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
            //    var dummy = WinAPI.LoadLibrary("kernel32.dll");
            // var fncptr = Marshal.GetFunctionPointerForDelegate( new WinAPI.HookProc(CBTHook));
            // var z = GCHandle.Alloc(new WinAPI.HookProc(CBTHook));
            // var ptr = z.AddrOfPinnedObject();
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
            switch ((WinAPI.HCBT)nCode)
            {
                case WinAPI.HCBT.MinMax:
                    if (wParam == outlookHwnd)
                    {
                        OutlookState = OutlookStateEnum.INBOX;
                    }
                    else
                    {
                        OutlookState = OutlookStateEnum.MINIMIZED;
                    }
                    break;
                case WinAPI.HCBT.MoveSize:
                    if (wParam == outlookHwnd)
                    {
                        var x = 1 + 1;
                    }
                    break;
                case WinAPI.HCBT.SetFocus:
                    break;
                default:
                    break;
            }

            return WinAPI.CallNextHookEx(_cbtHook, nCode, wParam, lParam);
        }
    }
}
