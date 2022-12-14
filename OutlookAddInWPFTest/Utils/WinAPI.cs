using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace OutlookAddInWPFTest.Utils
{
    public static class WinAPI
    {
        public delegate bool EnumChildWindowsCallback(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetThreadDpiAwarenessContext();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetAwarenessFromDpiAwarenessContext(IntPtr DPI_AWARENESS_CONTEXT);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern DPI_AWARENESS_CONTEXT SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT value);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern DPI_AWARENESS_CONTEXT SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT value);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool AreDpiAwarenessContextsEqual(DPI_AWARENESS_CONTEXT value1,
            DPI_AWARENESS_CONTEXT value2);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hwnd, int index);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int SetWindowLong(IntPtr hwnd, int index, int newStyle);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildWindowsCallback callback, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern int GetLastError();

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
        [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr SetWindowsHookEx(HookType hookType, [MarshalAs(UnmanagedType.FunctionPtr)] HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Ansi)] public static extern IntPtr LoadLibrary([MarshalAs(UnmanagedType.LPStr)] string lpFileName);

        public delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindow(IntPtr hWnd, GetWindowType uCmd);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ClipCursor(ref RECT rect);
        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(WinEvents eventMin, WinEvents eventMax, IntPtr
                hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
            uint idThread, WinEventFlags dwFlags);
        public delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
            IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);
        public enum WinEventFlags : uint
        {
            WINEVENT_OUTOFCONTEXT = 0x0000, // Events are ASYNC
            WINEVENT_SKIPOWNTHREAD = 0x0001, // Don't call back for events on installer's thread
            WINEVENT_SKIPOWNPROCESS = 0x0002, // Don't call back for events on installer's process
            WINEVENT_INCONTEXT = 0x0004, // Events are SYNC, this causes your dll to be injected into every process
        }

        public enum WinEvents : uint
        {
            /** The range of WinEvent constant values specified by the Accessibility Interoperability Alliance (AIA) for use across the industry.
        * For more information, see Allocation of WinEvent IDs. */
            EVENT_AIA_START = 0xA000,
            EVENT_AIA_END = 0xAFFF,

            /** The lowest and highest possible event values.
*/
            EVENT_MIN = 0x00000001,
            EVENT_MAX = 0x7FFFFFFF,

            /** An object's KeyboardShortcut property has changed. Server applications send this event for their accessible objects.
*/
            EVENT_OBJECT_ACCELERATORCHANGE = 0x8012,

            /** Sent when a window is cloaked. A cloaked window still exists, but is invisible to the user.
*/
            EVENT_OBJECT_CLOAKED = 0x8017,

            /** A window object's scrolling has ended. Unlike EVENT_SYSTEM_SCROLLEND, this event is associated with the scrolling window.
        * Whether the scrolling is horizontal or vertical scrolling, this event should be sent whenever the scroll action is completed. * The hwnd parameter of the WinEventProc callback function describes the scrolling window; the idObject parameter is OBJID_CLIENT, * and the idChild parameter is CHILDID_SELF. */
            EVENT_OBJECT_CONTENTSCROLLED = 0x8015,

            /** An object has been created. The system sends this event for the following user interface elements: caret, header control,
        * list-view control, tab control, toolbar control, tree view control, and window object. Server applications send this event * for their accessible objects. * Before sending the event for the parent object, servers must send it for all of an object's child objects. * Servers must ensure that all child objects are fully created and ready to accept IAccessible calls from clients before * the parent object sends this event. * Because a parent object is created after its child objects, clients must make sure that an object's parent has been created * before calling IAccessible::get_accParent, particularly if in-context hook functions are used. */
            EVENT_OBJECT_CREATE = 0x8000,

            /** An object's DefaultAction property has changed. The system sends this event for dialog boxes. Server applications send
        * this event for their accessible objects. */
            EVENT_OBJECT_DEFACTIONCHANGE = 0x8011,

            /** An object's Description property has changed. Server applications send this event for their accessible objects.
*/
            EVENT_OBJECT_DESCRIPTIONCHANGE = 0x800D,

            /** An object has been destroyed. The system sends this event for the following user interface elements: caret, header control,
        * list-view control, tab control, toolbar control, tree view control, and window object. Server applications send this event for * their accessible objects. * Clients assume that all of an object's children are destroyed when the parent object sends this event. * After receiving this event, clients do not call an object's IAccessible properties or methods. However, the interface pointer * must remain valid as long as there is a reference count on it (due to COM rules), but the UI element may no longer be present. * Further calls on the interface pointer may return failure errors; to prevent this, servers create proxy objects and monitor * their life spans. */
            EVENT_OBJECT_DESTROY = 0x8001,

            /** The user started to drag an element. The hwnd, idObject, and idChild parameters of the WinEventProc callback function
        * identify the object being dragged. */
            EVENT_OBJECT_DRAGSTART = 0x8021,

            /** The user has ended a drag operation before dropping the dragged element on a drop target. The hwnd, idObject, and idChild
        * parameters of the WinEventProc callback function identify the object being dragged. */
            EVENT_OBJECT_DRAGCANCEL = 0x8022,

            /** The user dropped an element on a drop target. The hwnd, idObject, and idChild parameters of the WinEventProc callback
        * function identify the object being dragged. */
            EVENT_OBJECT_DRAGCOMPLETE = 0x8023,

            /** The user dragged an element into a drop target's boundary. The hwnd, idObject, and idChild parameters of the WinEventProc
        * callback function identify the drop target. */
            EVENT_OBJECT_DRAGENTER = 0x8024,

            /** The user dragged an element out of a drop target's boundary. The hwnd, idObject, and idChild parameters of the WinEventProc
        * callback function identify the drop target. */
            EVENT_OBJECT_DRAGLEAVE = 0x8025,

            /** The user dropped an element on a drop target. The hwnd, idObject, and idChild parameters of the WinEventProc callback
        * function identify the drop target. */
            EVENT_OBJECT_DRAGDROPPED = 0x8026,

            /** The highest object event value.
*/
            EVENT_OBJECT_END = 0x80FF,

            /** An object has received the keyboard focus. The system sends this event for the following user interface elements:
        * list-view control, menu bar, pop-up menu, switch window, tab control, tree view control, and window object. * Server applications send this event for their accessible objects. * The hwnd parameter of the WinEventProc callback function identifies the window that receives the keyboard focus. */
            EVENT_OBJECT_FOCUS = 0x8005,

            /** An object's Help property has changed. Server applications send this event for their accessible objects.
*/
            EVENT_OBJECT_HELPCHANGE = 0x8010,

            /** An object is hidden. The system sends this event for the following user interface elements: caret and cursor.
        * Server applications send this event for their accessible objects. * When this event is generated for a parent object, all child objects are already hidden. * Server applications do not send this event for the child objects. * Hidden objects include the STATE_SYSTEM_INVISIBLE flag; shown objects do not include this flag. The EVENT_OBJECT_HIDE event * also indicates that the STATE_SYSTEM_INVISIBLE flag is set. Therefore, servers do not send the EVENT_STATE_CHANGE event in * this case. */
            EVENT_OBJECT_HIDE = 0x8003,

            /** A window that hosts other accessible objects has changed the hosted objects. A client might need to query the host
        * window to discover the new hosted objects, especially if the client has been monitoring events from the window. * A hosted object is an object from an accessibility framework (MSAA or UI Automation) that is different from that of the host. * Changes in hosted objects that are from the same framework as the host should be handed with the structural change events, * such as EVENT_OBJECT_CREATE for MSAA. For more info see comments within winuser.h. */
            EVENT_OBJECT_HOSTEDOBJECTSINVALIDATED = 0x8020,

            /** An IME window has become hidden.
*/
            EVENT_OBJECT_IME_HIDE = 0x8028,

            /** An IME window has become visible.
*/
            EVENT_OBJECT_IME_SHOW = 0x8027,

            /** The size or position of an IME window has changed.
*/
            EVENT_OBJECT_IME_CHANGE = 0x8029,

            /** An object has been invoked; for example, the user has clicked a button. This event is supported by common controls and is
        * used by UI Automation. * For this event, the hwnd, ID, and idChild parameters of the WinEventProc callback function identify the item that is invoked. */
            EVENT_OBJECT_INVOKED = 0x8013,

            /** An object that is part of a live region has changed. A live region is an area of an application that changes frequently
        * and/or asynchronously. */
            EVENT_OBJECT_LIVEREGIONCHANGED = 0x8019,

            /** An object has changed location, shape, or size. The system sends this event for the following user interface elements:
        * caret and window objects. Server applications send this event for their accessible objects. * This event is generated in response to a change in the top-level object within the object hierarchy; it is not generated for any * children that the object might have. For example, if the user resizes a window, the system sends this notification for the window, * but not for the menu bar, title bar, scroll bar, or other objects that have also changed. * The system does not send this event for every non-floating child window when the parent moves. However, if an application explicitly * resizes child windows as a result of resizing the parent window, the system sends multiple events for the resized children. * If an object's State property is set to STATE_SYSTEM_FLOATING, the server sends EVENT_OBJECT_LOCATIONCHANGE whenever the object changes * location. If an object does not have this state, servers only trigger this event when the object moves in relation to its parent. * For this event notification, the idChild parameter of the WinEventProc callback function identifies the child object that has changed. */
            EVENT_OBJECT_LOCATIONCHANGE = 0x800B,

            /** An object's Name property has changed. The system sends this event for the following user interface elements: check box,
        * cursor, list-view control, push button, radio button, status bar control, tree view control, and window object. Server * * applications send this event for their accessible objects. */
            EVENT_OBJECT_NAMECHANGE = 0x800C,

            /** An object has a new parent object. Server applications send this event for their accessible objects.
*/
            EVENT_OBJECT_PARENTCHANGE = 0x800F,

            /** A container object has added, removed, or reordered its children. The system sends this event for the following user
        * interface elements: header control, list-view control, toolbar control, and window object. Server applications send this * event as appropriate for their accessible objects. * For example, this event is generated by a list-view object when the number of child elements or the order of the elements changes. * This event is also sent by a parent window when the Z-order for the child windows changes. */
            EVENT_OBJECT_REORDER = 0x8004,

            /** The selection within a container object has changed. The system sends this event for the following user interface elements:
        * list-view control, tab control, tree view control, and window object. Server applications send this event for their accessible * objects. * This event signals a single selection: either a child is selected in a container that previously did not contain any selected children, * or the selection has changed from one child to another. * The hwnd and idObject parameters of the WinEventProc callback function describe the container; the idChild parameter identifies the object * that is selected. If the selected child is a window that also contains objects, the idChild parameter is OBJID_WINDOW. */
            EVENT_OBJECT_SELECTION = 0x8006,

            /** A child within a container object has been added to an existing selection. The system sends this event for the following user
        * interface elements: list box, list-view control, and tree view control. Server applications send this event for their accessible * objects. * The hwnd and idObject parameters of the WinEventProc callback function describe the container. The idChild parameter is the child that * is added to the selection. */
            EVENT_OBJECT_SELECTIONADD = 0x8007,

            /** An item within a container object has been removed from the selection. The system sends this event for the following user
        * interface elements: list box, list-view control, and tree view control. Server applications send this event for their accessible * objects. * This event signals that a child is removed from an existing selection. * The hwnd and idObject parameters of the WinEventProc callback function describe the container; the idChild parameter identifies * the child that has been removed from the selection. */
            EVENT_OBJECT_SELECTIONREMOVE = 0x8008,

            /** Numerous selection changes have occurred within a container object. The system sends this event for list boxes; server
        * applications send it for their accessible objects. * This event is sent when the selected items within a control have changed substantially. The event informs the client * that many selection changes have occurred, and it is sent instead of several * EVENT_OBJECT_SELECTIONADD or EVENT_OBJECT_SELECTIONREMOVE events. The client * queries for the selected items by calling the container object's IAccessible::get_accSelection method and * enumerating the selected items. For this event notification, the hwnd and idObject parameters of the WinEventProc callback * function describe the container in which the changes occurred. */
            EVENT_OBJECT_SELECTIONWITHIN = 0x8009,

            /** A hidden object is shown. The system sends this event for the following user interface elements: caret, cursor, and window
        * object. Server applications send this event for their accessible objects. * Clients assume that when this event is sent by a parent object, all child objects are already displayed. * Therefore, server applications do not send this event for the child objects. * Hidden objects include the STATE_SYSTEM_INVISIBLE flag; shown objects do not include this flag. * The EVENT_OBJECT_SHOW event also indicates that the STATE_SYSTEM_INVISIBLE flag is cleared. Therefore, servers * do not send the EVENT_STATE_CHANGE event in this case. */
            EVENT_OBJECT_SHOW = 0x8002,

            /** An object's state has changed. The system sends this event for the following user interface elements: check box, combo box,
        * header control, push button, radio button, scroll bar, toolbar control, tree view control, up-down control, and window object. * Server applications send this event for their accessible objects. * For example, a state change occurs when a button object is clicked or released, or when an object is enabled or disabled. * For this event notification, the idChild parameter of the WinEventProc callback function identifies the child object whose state has changed. */
            EVENT_OBJECT_STATECHANGE = 0x800A,

            /** The conversion target within an IME composition has changed. The conversion target is the subset of the IME composition
        * which is actively selected as the target for user-initiated conversions. */
            EVENT_OBJECT_TEXTEDIT_CONVERSIONTARGETCHANGED = 0x8030,

            /** An object's text selection has changed. This event is supported by common controls and is used by UI Automation.
        * The hwnd, ID, and idChild parameters of the WinEventProc callback function describe the item that is contained in the updated text selection. */
            EVENT_OBJECT_TEXTSELECTIONCHANGED = 0x8014,

            /** Sent when a window is uncloaked. A cloaked window still exists, but is invisible to the user.
*/
            EVENT_OBJECT_UNCLOAKED = 0x8018,

            /** An object's Value property has changed. The system sends this event for the user interface elements that include the scroll
        * bar and the following controls: edit, header, hot key, progress bar, slider, and up-down. Server applications send this event * for their accessible objects. */
            EVENT_OBJECT_VALUECHANGE = 0x800E,

            /** The range of event constant values reserved for OEMs. For more information, see Allocation of WinEvent IDs.
*/
            EVENT_OEM_DEFINED_START = 0x0101,
            EVENT_OEM_DEFINED_END = 0x01FF,

            /** An alert has been generated. Server applications should not send this event.
*/
            EVENT_SYSTEM_ALERT = 0x0002,

            /** A preview rectangle is being displayed.
*/
            EVENT_SYSTEM_ARRANGMENTPREVIEW = 0x8016,

            /** A window has lost mouse capture. This event is sent by the system, never by servers.
*/
            EVENT_SYSTEM_CAPTUREEND = 0x0009,

            /** A window has received mouse capture. This event is sent by the system, never by servers.
*/
            EVENT_SYSTEM_CAPTURESTART = 0x0008,

            /** A window has exited context-sensitive Help mode. This event is not sent consistently by the system.
*/
            EVENT_SYSTEM_CONTEXTHELPEND = 0x000D,

            /** A window has entered context-sensitive Help mode. This event is not sent consistently by the system.
*/
            EVENT_SYSTEM_CONTEXTHELPSTART = 0x000C,

            /** The active desktop has been switched.
*/
            EVENT_SYSTEM_DESKTOPSWITCH = 0x0020,

            /** A dialog box has been closed. The system sends this event for standard dialog boxes; servers send it for custom dialog boxes.
        * This event is not sent consistently by the system. */
            EVENT_SYSTEM_DIALOGEND = 0x0011,

            /** A dialog box has been displayed. The system sends this event for standard dialog boxes, which are created using resource
        * templates or Win32 dialog box functions. Servers send this event for custom dialog boxes, which are windows that function as * dialog boxes but are not created in the standard way. * This event is not sent consistently by the system. */
            EVENT_SYSTEM_DIALOGSTART = 0x0010,

            /** An application is about to exit drag-and-drop mode. Applications that support drag-and-drop operations must send this event;
        * the system does not send this event. */
            EVENT_SYSTEM_DRAGDROPEND = 0x000F,

            /** An application is about to enter drag-and-drop mode. Applications that support drag-and-drop operations must send this
        * event because the system does not send it. */
            EVENT_SYSTEM_DRAGDROPSTART = 0x000E,

            /** The highest system event value.
*/
            EVENT_SYSTEM_END = 0x00FF,

            /** The foreground window has changed. The system sends this event even if the foreground window has changed to another window
        * in the same thread. Server applications never send this event. * For this event, the WinEventProc callback function's hwnd parameter is the handle to the window that is in the * foreground, the idObject parameter is OBJID_WINDOW, and the idChild parameter is CHILDID_SELF. */
            EVENT_SYSTEM_FOREGROUND = 0x0003,

            /** A pop-up menu has been closed. The system sends this event for standard menus; servers send it for custom menus.
        * When a pop-up menu is closed, the client receives this message, and then the EVENT_SYSTEM_MENUEND event. * This event is not sent consistently by the system. */
            EVENT_SYSTEM_MENUPOPUPEND = 0x0007,

            /** A pop-up menu has been displayed. The system sends this event for standard menus, which are identified by HMENU, and are
        * created using menu-template resources or Win32 menu functions. Servers send this event for custom menus, which are user * interface elements that function as menus but are not created in the standard way. This event is not sent consistently by the system. */
            EVENT_SYSTEM_MENUPOPUPSTART = 0x0006,

            /** A menu from the menu bar has been closed. The system sends this event for standard menus; servers send it for custom menus.
        * For this event, the WinEventProc callback function's hwnd, idObject, and idChild parameters refer to the control * that contains the menu bar or the control that activates the context menu. The hwnd parameter is the handle to the window * that is related to the event. The idObject parameter is OBJID_MENU or OBJID_SYSMENU for a menu, or OBJID_WINDOW for a * pop-up menu. The idChild parameter is CHILDID_SELF. */
            EVENT_SYSTEM_MENUEND = 0x0005,

            /** A menu item on the menu bar has been selected. The system sends this event for standard menus, which are identified
        * by HMENU, created using menu-template resources or Win32 menu API elements. Servers send this event for custom menus, * which are user interface elements that function as menus but are not created in the standard way. * For this event, the WinEventProc callback function's hwnd, idObject, and idChild parameters refer to the control * that contains the menu bar or the control that activates the context menu. The hwnd parameter is the handle to the window * related to the event. The idObject parameter is OBJID_MENU or OBJID_SYSMENU for a menu, or OBJID_WINDOW for a pop-up menu. * The idChild parameter is CHILDID_SELF.The system triggers more than one EVENT_SYSTEM_MENUSTART event that does not always * correspond with the EVENT_SYSTEM_MENUEND event. */
            EVENT_SYSTEM_MENUSTART = 0x0004,

            /** A window object is about to be restored. This event is sent by the system, never by servers.
*/
            EVENT_SYSTEM_MINIMIZEEND = 0x0017,

            /** A window object is about to be minimized. This event is sent by the system, never by servers.
*/
            EVENT_SYSTEM_MINIMIZESTART = 0x0016,

            /** The movement or resizing of a window has finished. This event is sent by the system, never by servers.
*/
            EVENT_SYSTEM_MOVESIZEEND = 0x000B,

            /** A window is being moved or resized. This event is sent by the system, never by servers.
*/
            EVENT_SYSTEM_MOVESIZESTART = 0x000A,

            /** Scrolling has ended on a scroll bar. This event is sent by the system for standard scroll bar controls and for
        * scroll bars that are attached to a window. Servers send this event for custom scroll bars, which are user interface * elements that function as scroll bars but are not created in the standard way. * The idObject parameter that is sent to the WinEventProc callback function is OBJID_HSCROLL for horizontal scroll bars, and * OBJID_VSCROLL for vertical scroll bars. */
            EVENT_SYSTEM_SCROLLINGEND = 0x0013,

            /** Scrolling has started on a scroll bar. The system sends this event for standard scroll bar controls and for scroll
        * bars attached to a window. Servers send this event for custom scroll bars, which are user interface elements that * function as scroll bars but are not created in the standard way. * The idObject parameter that is sent to the WinEventProc callback function is OBJID_HSCROLL for horizontal scrolls bars, * and OBJID_VSCROLL for vertical scroll bars. */
            EVENT_SYSTEM_SCROLLINGSTART = 0x0012,

            /** A sound has been played. The system sends this event when a system sound, such as one for a menu,
        * is played even if no sound is audible (for example, due to the lack of a sound file or a sound card). * Servers send this event whenever a custom UI element generates a sound. * For this event, the WinEventProc callback function receives the OBJID_SOUND value as the idObject parameter. */
            EVENT_SYSTEM_SOUND = 0x0001,

            /** The user has released ALT+TAB. This event is sent by the system, never by servers.
        * The hwnd parameter of the WinEventProc callback function identifies the window to which the user has switched. * If only one application is running when the user presses ALT+TAB, the system sends this event without a corresponding * EVENT_SYSTEM_SWITCHSTART event. */
            EVENT_SYSTEM_SWITCHEND = 0x0015,

            /** The user has pressed ALT+TAB, which activates the switch window. This event is sent by the system, never by servers.
        * The hwnd parameter of the WinEventProc callback function identifies the window to which the user is switching. * If only one application is running when the user presses ALT+TAB, the system sends an EVENT_SYSTEM_SWITCHEND event without a * corresponding EVENT_SYSTEM_SWITCHSTART event. */
            EVENT_SYSTEM_SWITCHSTART = 0x0014,

            /** The range of event constant values reserved for UI Automation event identifiers. For more information,
        * see Allocation of WinEvent IDs. */
            EVENT_UIA_EVENTID_START = 0x4E00,
            EVENT_UIA_EVENTID_END = 0x4EFF,

            /**
        * The range of event constant values reserved for UI Automation property-changed event identifiers. * For more information, see Allocation of WinEvent IDs. */
            EVENT_UIA_PROPID_START = 0x7500,
            EVENT_UIA_PROPID_END = 0x75FF
        }
        public enum GetWindowType : uint
        {
            /// <summary>
            /// The retrieved handle identifies the window of the same type that is highest in the Z order.
            /// <para/>
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDFIRST = 0,
            /// <summary>
            /// The retrieved handle identifies the window of the same type that is lowest in the Z order.
            /// <para />
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDLAST = 1,
            /// <summary>
            /// The retrieved handle identifies the window below the specified window in the Z order.
            /// <para />
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDNEXT = 2,
            /// <summary>
            /// The retrieved handle identifies the window above the specified window in the Z order.
            /// <para />
            /// If the specified window is a topmost window, the handle identifies a topmost window.
            /// If the specified window is a top-level window, the handle identifies a top-level window.
            /// If the specified window is a child window, the handle identifies a sibling window.
            /// </summary>
            GW_HWNDPREV = 3,
            /// <summary>
            /// The retrieved handle identifies the specified window's owner window, if any.
            /// </summary>
            GW_OWNER = 4,
            /// <summary>
            /// The retrieved handle identifies the child window at the top of the Z order,
            /// if the specified window is a parent window; otherwise, the retrieved handle is NULL.
            /// The function examines only child windows of the specified window. It does not examine descendant windows.
            /// </summary>
            GW_CHILD = 5,
            /// <summary>
            /// The retrieved handle identifies the enabled popup window owned by the specified window (the
            /// search uses the first such window found using GW_HWNDNEXT); otherwise, if there are no enabled
            /// popup windows, the retrieved handle is that of the specified window.
            /// </summary>
            GW_ENABLEDPOPUP = 6
        }

        public enum HCBT : int
        {
            MoveSize = 0,
            MinMax = 1,
            QueueSync = 2,
            CreateWnd = 3,
            DestroyWnd = 4,
            Activate = 5,
            ClickSkipped = 6,
            KeySkipped = 7,
            SysCommand = 8,
            SetFocus = 9
        }

        /// <summary>
        /// Enumerates the valid hook types passed as the idHook parameter into a call to SetWindowsHookEx.
        /// </summary>
        public enum HookType : int
        {
            /// <summary>
            /// Installs a hook procedure that monitors messages generated as a result of an input event in a dialog box,
            /// message box, menu, or scroll bar. For more information, see the MessageProc hook procedure.
            /// </summary>
            WH_MSGFILTER = -1,
            /// <summary>
            /// Installs a hook procedure that records input messages posted to the system message queue. This hook is
            /// useful for recording macros. For more information, see the JournalRecordProc hook procedure.
            /// </summary>
            WH_JOURNALRECORD = 0,
            /// <summary>
            /// Installs a hook procedure that posts messages previously recorded by a WH_JOURNALRECORD hook procedure.
            /// For more information, see the JournalPlaybackProc hook procedure.
            /// </summary>
            WH_JOURNALPLAYBACK = 1,
            /// <summary>
            /// Installs a hook procedure that monitors keystroke messages. For more information, see the KeyboardProc
            /// hook procedure.
            /// </summary>
            WH_KEYBOARD = 2,
            /// <summary>
            /// Installs a hook procedure that monitors messages posted to a message queue. For more information, see the
            /// GetMsgProc hook procedure.
            /// </summary>
            WH_GETMESSAGE = 3,
            /// <summary>
            /// Installs a hook procedure that monitors messages before the system sends them to the destination window
            /// procedure. For more information, see the CallWndProc hook procedure.
            /// </summary>
            WH_CALLWNDPROC = 4,
            /// <summary>
            /// Installs a hook procedure that receives notifications useful to a CBT application. For more information,
            /// see the CBTProc hook procedure.
            /// </summary>
            WH_CBT = 5,
            /// <summary>
            /// Installs a hook procedure that monitors messages generated as a result of an input event in a dialog box,
            /// message box, menu, or scroll bar. The hook procedure monitors these messages for all applications in the
            /// same desktop as the calling thread. For more information, see the SysMsgProc hook procedure.
            /// </summary>
            WH_SYSMSGFILTER = 6,
            /// <summary>
            /// Installs a hook procedure that monitors mouse messages. For more information, see the MouseProc hook
            /// procedure.
            /// </summary>
            WH_MOUSE = 7,
            /// <summary>
            ///
            /// </summary>
            WH_HARDWARE = 8,
            /// <summary>
            /// Installs a hook procedure useful for debugging other hook procedures. For more information, see the
            /// DebugProc hook procedure.
            /// </summary>
            WH_DEBUG = 9,
            /// <summary>
            /// Installs a hook procedure that receives notifications useful to shell applications. For more information,
            /// see the ShellProc hook procedure.
            /// </summary>
            WH_SHELL = 10,
            /// <summary>
            /// Installs a hook procedure that will be called when the application's foreground thread is about to become
            /// idle. This hook is useful for performing low priority tasks during idle time. For more information, see the
            /// ForegroundIdleProc hook procedure.
            /// </summary>
            WH_FOREGROUNDIDLE = 11,
            /// <summary>
            /// Installs a hook procedure that monitors messages after they have been processed by the destination window
            /// procedure. For more information, see the CallWndRetProc hook procedure.
            /// </summary>
            WH_CALLWNDPROCRET = 12,
            /// <summary>
            /// Installs a hook procedure that monitors low-level keyboard input events. For more information, see the
            /// LowLevelKeyboardProc hook procedure.
            /// </summary>
            WH_KEYBOARD_LL = 13,
            /// <summary>
            /// Installs a hook procedure that monitors low-level mouse input events. For more information, see the
            /// LowLevelMouseProc hook procedure.
            /// </summary>
            WH_MOUSE_LL = 14
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct CBTACTIVATESTRUCT
        {
            public bool fMouse;
            public IntPtr hWndActive;
        }

        public struct DPI_AWARENESS_CONTEXT
        {
            private IntPtr value;

            private DPI_AWARENESS_CONTEXT(IntPtr value)
            {
                this.value = value;
            }

            public static implicit operator DPI_AWARENESS_CONTEXT(IntPtr value)
            {
                return new DPI_AWARENESS_CONTEXT(value);
            }

            public static implicit operator IntPtr(DPI_AWARENESS_CONTEXT context)
            {
                return context.value;
            }

            public static bool operator ==(IntPtr context1, DPI_AWARENESS_CONTEXT context2)
            {
                return AreDpiAwarenessContextsEqual(context1, context2);
            }

            public static bool operator !=(IntPtr context1, DPI_AWARENESS_CONTEXT context2)
            {
                return !AreDpiAwarenessContextsEqual(context1, context2);
            }

            public override bool Equals(object obj)
            {
                return base.Equals(obj);
            }

            public override int GetHashCode()
            {
                return base.GetHashCode();
            }
        }

        private static DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_HANDLE = IntPtr.Zero;

        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_INVALID = IntPtr.Zero;
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_UNAWARE = new IntPtr(-1);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_SYSTEM_AWARE = new IntPtr(-2);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE = new IntPtr(-3);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
        public static readonly DPI_AWARENESS_CONTEXT DPI_AWARENESS_CONTEXT_UNAWARE_GDISCALED = new IntPtr(-5);

        public static DPI_AWARENESS_CONTEXT[] DpiAwarenessContexts =
        {
            DPI_AWARENESS_CONTEXT_UNAWARE,
            DPI_AWARENESS_CONTEXT_SYSTEM_AWARE,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE,
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2,
            DPI_AWARENESS_CONTEXT_UNAWARE_GDISCALED,
        };

        public const int WS_EX_TRANSPARENT = 0x00000020;
        public const int GWL_EXSTYLE = (-20);

        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static IntPtr FindChildWindowByClassName(IntPtr hwndParent, string className,
            bool findVisibleOnly = true)
        {
            var res = IntPtr.Zero;
            var cls = new StringBuilder(className.Length + 5);

            EnumChildWindows(hwndParent, delegate(IntPtr hwndChild, IntPtr lParam)
            {
                GetClassName(hwndChild, cls, cls.Capacity);
                var flag = IsWindowVisible(hwndChild);
                if ((!findVisibleOnly || flag) && cls.ToString() == className)
                {
                    res = hwndChild;
                    return false;
                }

                return true;
            }, IntPtr.Zero);
            return res;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public static implicit operator System.Drawing.Point(POINT p)
            {
                return new System.Drawing.Point(p.X, p.Y);
            }

            public static implicit operator POINT(System.Drawing.Point p)
            {
                return new POINT(p.X, p.Y);
            }
        }

        public enum ShowWindowCommands
        {
            /// <summary>
            /// Hides the window and activates another window.
            /// </summary>
            Hide = 0,

            /// <summary>
            /// Activates and displays a window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when displaying the window
            /// for the first time.
            /// </summary>
            Normal = 1,

            /// <summary>
            /// Activates the window and displays it as a minimized window.
            /// </summary>
            ShowMinimized = 2,

            /// <summary>
            /// Maximizes the specified window.
            /// </summary>
            Maximize = 3, // is this the right value?

            /// <summary>
            /// Activates the window and displays it as a maximized window.
            /// </summary>      
            ShowMaximized = 3,

            /// <summary>
            /// Displays a window in its most recent size and position. This value
            /// is similar to <see cref="Win32.ShowWindowCommand.Normal"/>, except
            /// the window is not activated.
            /// </summary>
            ShowNoActivate = 4,

            /// <summary>
            /// Activates the window and displays it in its current size and position.
            /// </summary>
            Show = 5,

            /// <summary>
            /// Minimizes the specified window and activates the next top-level
            /// window in the Z order.
            /// </summary>
            Minimize = 6,

            /// <summary>
            /// Displays the window as a minimized window. This value is similar to
            /// <see cref="Win32.ShowWindowCommand.ShowMinimized"/>, except the
            /// window is not activated.
            /// </summary>
            ShowMinNoActive = 7,

            /// <summary>
            /// Displays the window in its current size and position. This value is
            /// similar to <see cref="Win32.ShowWindowCommand.Show"/>, except the
            /// window is not activated.
            /// </summary>
            ShowNA = 8,

            /// <summary>
            /// Activates and displays the window. If the window is minimized or
            /// maximized, the system restores it to its original size and position.
            /// An application should specify this flag when restoring a minimized window.
            /// </summary>
            Restore = 9,

            /// <summary>
            /// Sets the show state based on the SW_* value specified in the
            /// STARTUPINFO structure passed to the CreateProcess function by the
            /// program that started the application.
            /// </summary>
            ShowDefault = 10,

            /// <summary>
            ///  <b>Windows 2000/XP:</b> Minimizes a window, even if the thread
            /// that owns the window is not responding. This flag should only be
            /// used when minimizing windows from a different thread.
            /// </summary>
            ForceMinimize = 11
        }

        /// <summary>
        /// Contains information about the placement of a window on the screen.
        /// </summary>
        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            /// <summary>
            /// The length of the structure, in bytes. Before calling the GetWindowPlacement or SetWindowPlacement functions, set this member to sizeof(WINDOWPLACEMENT).
            /// <para>
            /// GetWindowPlacement and SetWindowPlacement fail if this member is not set correctly.
            /// </para>
            /// </summary>
            public int Length;

            /// <summary>
            /// Specifies flags that control the position of the minimized window and the method by which the window is restored.
            /// </summary>
            public int Flags;

            /// <summary>
            /// The current show state of the window.
            /// </summary>
            public ShowWindowCommands ShowCmd;

            /// <summary>
            /// The coordinates of the window's upper-left corner when the window is minimized.
            /// </summary>
            public POINT MinPosition;

            /// <summary>
            /// The coordinates of the window's upper-left corner when the window is maximized.
            /// </summary>
            public POINT MaxPosition;

            /// <summary>
            /// The window's coordinates when the window is in the restored position.
            /// </summary>
            public RECT NormalPosition;

            /// <summary>
            /// Gets the default (empty) value.
            /// </summary>
            public static WINDOWPLACEMENT Default
            {
                get
                {
                    WINDOWPLACEMENT result = new WINDOWPLACEMENT();
                    result.Length = Marshal.SizeOf(result);
                    return result;
                }
            }

        }
        private static int MAKELPARAM(int p, int p_2)
        {
            return ((p_2 << 16) | (p & 0xFFFF));
        }
        public static void ClickMouseButton(IntPtr hwnd, bool isLeft, POINT screenPoint)
        {
            ScreenToClient(hwnd, ref screenPoint);
            var messageDown = (uint)(isLeft ? 0x0201 : 0x0204);
            var messageUp = (uint)(isLeft ? 0x0202 : 0x0205);
            var res1 = WinAPI.SendMessage(hwnd, messageDown, 1, MAKELPARAM(screenPoint.X, screenPoint.Y));
            var err = GetLastError();
            var res2 = WinAPI.SendMessage(hwnd, messageUp, 0, MAKELPARAM(screenPoint.X, screenPoint.Y));
        }
    }
}
