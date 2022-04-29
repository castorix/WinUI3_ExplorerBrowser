using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

// https://app.box.com/s/wmtnw883g6qjvtzuwzxszv8y7x84xj33

namespace WinUI3_ExplorerBrowser
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {

        public enum HRESULT : int
        {
            S_OK = 0,
            S_FALSE = 1,
            E_NOTIMPL = unchecked((int)0x80004001),
            E_NOINTERFACE = unchecked((int)0x80004002),
            E_POINTER = unchecked((int)0x80004003),
            E_FAIL = unchecked((int)0x80004005),
            E_UNEXPECTED = unchecked((int)0x8000FFFF),
            E_OUTOFMEMORY = unchecked((int)0x8007000E),
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
            public RECT(int left, int top, int right, int bottom)
            {
                this.left = left;
                this.top = top;
                this.right = right;
                this.bottom = bottom;
            }
        }

        [DllImport("User32.dll", SetLastError = true)]
        public static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("User32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        public const int SWP_NOSIZE = 0x0001;
        public const int SWP_NOMOVE = 0x0002;
        public const int SWP_NOZORDER = 0x0004;
        public const int SWP_NOREDRAW = 0x0008;
        public const int SWP_NOACTIVATE = 0x0010;
        public const int SWP_FRAMECHANGED = 0x0020;  /* The frame changed: send WM_NCCALCSIZE */
        public const int SWP_SHOWWINDOW = 0x0040;
        public const int SWP_HIDEWINDOW = 0x0080;
        public const int SWP_NOCOPYBITS = 0x0100;
        public const int SWP_NOOWNERZORDER = 0x0200;  /* Don't do owner Z ordering */
        public const int SWP_NOSENDCHANGING = 0x0400;  /* Don't send WM_WINDOWPOSCHANGING */
        public const int SWP_DRAWFRAME = SWP_FRAMECHANGED;
        public const int SWP_NOREPOSITION = SWP_NOOWNERZORDER;
        public const int SWP_DEFERERASE = 0x2000;
        public const int SWP_ASYNCWINDOWPOS = 0x4000;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [ComImport]
        [Guid("dfd3b6b5-c10c-4be9-85f6-a66969f402f6")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IExplorerBrowser
        {
            HRESULT Initialize(IntPtr hwndParent, ref RECT prc, ref FOLDERSETTINGS pfs);
            HRESULT Destroy();
            HRESULT SetRect(IntPtr phdwp, ref RECT rcBrowser);
            HRESULT SetPropertyBag(string pszPropertyBag);
            HRESULT SetEmptyText(string pszEmptyText);
            HRESULT SetFolderSettings(ref FOLDERSETTINGS pfs);

            HRESULT Advise(IntPtr psbe, out int pdwCookie);
            //HRESULT Advise([In()] [MarshalAs(UnmanagedType.IUnknown)] object punk, out int pdwCookie);
            //HRESULT Advise(IExplorerBrowserEvents punk, out int pdwCookie);

            HRESULT Unadvise(int dwCookie);
            HRESULT SetOptions(EXPLORER_BROWSER_OPTIONS dwFlag);
            HRESULT GetOptions(ref EXPLORER_BROWSER_OPTIONS pdwFlag);
            HRESULT BrowseToIDList(IntPtr pidl, uint uFlags);

            // IUnknown *punk,
            HRESULT BrowseToObject(IntPtr punk, uint uFlags);
            HRESULT FillFromObject(IntPtr punk, EXPLORER_BROWSER_FILL_FLAGS dwFlags);
            HRESULT RemoveAll();
            HRESULT GetCurrentView(ref Guid riid, ref IntPtr ppv);
        }

        public const int SBSP_ABSOLUTE = 0x0;

        public enum EXPLORER_BROWSER_OPTIONS : int
        {
            EBO_NONE = 0,
            EBO_NAVIGATEONCE = 0x1,
            EBO_SHOWFRAMES = 0x2,
            EBO_ALWAYSNAVIGATE = 0x4,
            EBO_NOTRAVELLOG = 0x8,
            EBO_NOWRAPPERWINDOW = 0x10,
            EBO_HTMLSHAREPOINTVIEW = 0x20,
            EBO_NOBORDER = 0x40,
            EBO_NOPERSISTVIEWSTATE = 0x80
        }

        public enum EXPLORER_BROWSER_FILL_FLAGS : int
        {
            EBF_NONE = 0,
            EBF_SELECTFROMDATAOBJECT = 0x100,
            EBF_NODROPTARGET = 0x200
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct FOLDERSETTINGS
        {
            public int ViewMode;
            public uint fFlags;
        }

        public enum FOLDERVIEWMODE : int
        {
            FVM_AUTO = -1,
            FVM_FIRST = 1,
            FVM_ICON = 1,
            FVM_SMALLICON = 2,
            FVM_LIST = 3,
            FVM_DETAILS = 4,
            FVM_THUMBNAIL = 5,
            FVM_TILE = 6,
            FVM_THUMBSTRIP = 7,
            FVM_CONTENT = 8,
            FVM_LAST = 8
        }

        public enum FOLDERFLAGS
        {
            FWF_NONE = 0,
            FWF_AUTOARRANGE = 0x1,
            FWF_ABBREVIATEDNAMES = 0x2,
            FWF_SNAPTOGRID = 0x4,
            FWF_OWNERDATA = 0x8,
            FWF_BESTFITWINDOW = 0x10,
            FWF_DESKTOP = 0x20,
            FWF_SINGLESEL = 0x40,
            FWF_NOSUBFOLDERS = 0x80,
            FWF_TRANSPARENT = 0x100,
            FWF_NOCLIENTEDGE = 0x200,
            FWF_NOSCROLL = 0x400,
            FWF_ALIGNLEFT = 0x800,
            FWF_NOICONS = 0x1000,
            FWF_SHOWSELALWAYS = 0x2000,
            FWF_NOVISIBLE = 0x4000,
            FWF_SINGLECLICKACTIVATE = 0x8000,
            FWF_NOWEBVIEW = 0x10000,
            FWF_HIDEFILENAMES = 0x20000,
            FWF_CHECKSELECT = 0x40000,
            FWF_NOENUMREFRESH = 0x80000,
            FWF_NOGROUPING = 0x100000,
            FWF_FULLROWSELECT = 0x200000,
            FWF_NOFILTERS = 0x400000,
            FWF_NOCOLUMNHEADER = 0x800000,
            FWF_NOHEADERINALLVIEWS = 0x1000000,
            FWF_EXTENDEDTILES = 0x2000000,
            FWF_TRICHECKSELECT = 0x4000000,
            FWF_AUTOCHECKSELECT = 0x8000000,
            FWF_NOBROWSERVIEWSTATE = 0x10000000,
            FWF_SUBSETGROUPS = 0x20000000,
            FWF_USESEARCHFOLDER = 0x40000000,
            FWF_ALLOWRTLREADING = unchecked((int)0x80000000)
        }

        [ComImport]
        [Guid("e07010ec-bc17-44c0-97b0-46c7c95b9edc")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IExplorerPaneVisibility
        {
            [PreserveSig]
            HRESULT GetPaneState(ref Guid ep, out EXPLORERPANESTATE peps);
        }

        public enum EXPLORERPANESTATE
        {
            EPS_DONTCARE = 0,
            EPS_DEFAULT_ON = 0x1,
            EPS_DEFAULT_OFF = 0x2,
            EPS_STATEMASK = 0xFFFF,
            EPS_INITIALSTATE = 0x10000,
            EPS_FORCE = 0x20000
        }

        private class CExplorerBrowserHost : CExplorerBrowserHost.IServiceProvider, IExplorerPaneVisibility, IExplorerBrowserEvents
        {
            [ComImport()]
            [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
            [Guid("6d5140c1-7436-11ce-8034-00aa006009fa")]
            public interface IServiceProvider
            {
                [PreserveSig()]
                HRESULT QueryService(ref Guid guidService, ref Guid riid, ref IntPtr ppvObject);
            }

            [DllImport("Shlwapi.DLL", SetLastError = true, CharSet = CharSet.Unicode)]
            internal static extern HRESULT IUnknown_SetSite([In()][MarshalAs(UnmanagedType.IUnknown)] object punk, [In()][MarshalAs(UnmanagedType.IUnknown)] object punkSite);

            public HRESULT QueryService(ref Guid guidService, ref Guid riid, ref IntPtr ppvObject)
            {
                HRESULT hr = HRESULT.S_OK;
                if (guidService.CompareTo(new Guid("e07010ec-bc17-44c0-97b0-46c7c95b9edc")) == 0)
                {
                    ppvObject = Marshal.GetComInterfaceForObject(this, typeof(IExplorerPaneVisibility));
                    hr = HRESULT.S_OK;
                }
                else
                {
                    IntPtr nullObj = IntPtr.Zero;
                    ppvObject = nullObj;
                    hr = HRESULT.E_NOINTERFACE;
                }
                return hr;
            }

            public HRESULT GetPaneState(ref Guid explorerPane, out EXPLORERPANESTATE peps)
            {
                // Test force all on ON 
                // peps = EXPLORERPANESTATE.EPS_DEFAULT_ON Or EXPLORERPANESTATE.EPS_INITIALSTATE Or EXPLORERPANESTATE.EPS_FORCE
                peps = EXPLORERPANESTATE.EPS_DEFAULT_ON | EXPLORERPANESTATE.EPS_INITIALSTATE;
                return HRESULT.S_OK;
            }

            public IExplorerBrowser pExplorerBrowser = null;

            private Guid EP_NavPane = new Guid("{cb316b22-25f7-42b8-8a09-540d23a43c2f}");
            private Guid EP_Commands = new Guid("{d9745868-ca5f-4a76-91cd-f5a129fbb076}");
            private Guid EP_Commands_Organize = new Guid("{72e81700-e3ec-4660-bf24-3c3b7b648806}");
            private Guid EP_Commands_View = new Guid("{21f7c32d-eeaa-439b-bb51-37b96fd6a943}");
            private Guid EP_DetailsPane = new Guid("{43abf98b-89b8-472d-b9ce-e69b8229f019}");
            private Guid EP_PreviewPane = new Guid("{893c63d1-45c8-4d17-be19-223be71be365}");
            private Guid EP_QueryPane = new Guid("{65bcde4f-4f07-4f27-83a7-1afca4df7ddd}");
            private Guid EP_AdvQueryPane = new Guid("{b4e9db8b-34ba-4c39-b5cc-16a1bd2c411c}");
            private Guid EP_StatusBar = new Guid("{65fe56ce-5cfe-4bc4-ad8a-7ae3fe7e8f7c}");
            private Guid EP_Ribbon = new Guid("{D27524A8-C9F2-4834-A106-DF8889FD4F37}");

            private int dwEBAdvise;
            private MainWindow _mw;

            public CExplorerBrowserHost(MainWindow mw, IntPtr hWnd, RECT rc)
            {
                _mw = mw;
                Guid CLSID_ExplorerBrowser = new Guid("71f96385-ddd6-48d3-a0c1-ae06e8b055fb");
                Type ExplorerBrowserType = Type.GetTypeFromCLSID(CLSID_ExplorerBrowser, true);
                object ExplorerBrowser = Activator.CreateInstance(ExplorerBrowserType);
                pExplorerBrowser = (IExplorerBrowser)ExplorerBrowser;
                FOLDERSETTINGS fs;
                // fs.ViewMode = FOLDERVIEWMODE.FVM_THUMBNAIL
                fs.ViewMode = (int)FOLDERVIEWMODE.FVM_DETAILS;
                fs.fFlags = (int)FOLDERFLAGS.FWF_HIDEFILENAMES;
                //RECT rc;
                //GetClientRect(hWnd, out rc);
                //rc.right = 600;
                //rc.bottom = 400;
                if ((pExplorerBrowser != null))
                {
                    HRESULT hr = pExplorerBrowser.Initialize(hWnd, rc, fs);
                    if ((hr == HRESULT.S_OK))
                    {
                        pExplorerBrowser.SetOptions(EXPLORER_BROWSER_OPTIONS.EBO_SHOWFRAMES | EXPLORER_BROWSER_OPTIONS.EBO_ALWAYSNAVIGATE | EXPLORER_BROWSER_OPTIONS.EBO_NOTRAVELLOG | EXPLORER_BROWSER_OPTIONS.EBO_NOWRAPPERWINDOW | EXPLORER_BROWSER_OPTIONS.EBO_HTMLSHAREPOINTVIEW | EXPLORER_BROWSER_OPTIONS.EBO_NOBORDER | EXPLORER_BROWSER_OPTIONS.EBO_NOPERSISTVIEWSTATE);

                        //IntPtr pUnknown = Marshal.GetIUnknownForObject(pExplorerBrowser);
                        //IntPtr pUnknownSite = Marshal.GetIUnknownForObject((IServiceProvider)this);
                        //hr = IUnknown_SetSite(pUnknown, pUnknownSite)
                        //hr = IUnknown_SetSite(pExplorerBrowser, this);

                        hr = pExplorerBrowser.Advise(Marshal.GetComInterfaceForObject(this, typeof(IExplorerBrowserEvents)), out dwEBAdvise);
                        //hr = pExplorerBrowser.Advise(this, out _dwEBAdvise);
                        hr = IUnknown_SetSite(pExplorerBrowser, Marshal.GetComInterfaceForObject(this, typeof(IServiceProvider)));

                        IntPtr pidlInit = IntPtr.Zero;
                        Guid FOLDERID_ComputerFolder = new Guid("0AC0837C-BBF8-452A-850D-79D08E667CA7");
                        hr = SHGetKnownFolderIDList(ref FOLDERID_ComputerFolder, 0, IntPtr.Zero, ref pidlInit);
                        if ((hr == HRESULT.S_OK))
                            pExplorerBrowser.BrowseToIDList(pidlInit, SBSP_ABSOLUTE);
                        // Set DarkMode on right pane
                        //uint rgflnOut = 0;
                        //hr = SHILCreateFromPath("C:", ref pidlInit, ref rgflnOut);
                        //if (hr == HRESULT.S_OK)
                        //{
                        //    pExplorerBrowser.BrowseToIDList(pidlInit, SBSP_ABSOLUTE);
                        //}                      
                    }
                }
            }

            HRESULT IExplorerBrowserEvents.OnNavigationPending(IntPtr pidlFolder)
            {
                return HRESULT.S_OK;
            }

            HRESULT IExplorerBrowserEvents.OnViewCreated(IShellView psv)
            //HRESULT IExplorerBrowserEvents.OnViewCreated([MarshalAs(UnmanagedType.IUnknown)] object psv)
            {
                return HRESULT.S_OK;
            }

            HRESULT IExplorerBrowserEvents.OnNavigationComplete(IntPtr pidlFolder)
            {  
                _mw.OnNavigationComplete(pidlFolder);
                return HRESULT.S_OK;
            }

            HRESULT IExplorerBrowserEvents.OnNavigationFailed(IntPtr pidlFolder)
            {
                return HRESULT.S_OK;
            }
        }

        [ComImport, Guid("cde725b0-ccc9-4519-917e-325d72fab4ce"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IFolderView
        {
            HRESULT GetCurrentViewMode([Out] out uint pViewMode);
            HRESULT SetCurrentViewMode(uint ViewMode);
            HRESULT GetFolder(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppv);
            HRESULT Item(int iItemIndex, out IntPtr ppidl);
            HRESULT ItemCount(uint uFlags, out int pcItems);
            [PreserveSig()]
            HRESULT Items(uint uFlags, ref Guid riid, [Out, MarshalAs(UnmanagedType.IUnknown)] out object ppv);
            HRESULT GetSelectionMarkedItem(out int piItem);
            HRESULT GetFocusedItem(out int piItem);
            HRESULT GetItemPosition(IntPtr pidl, out Point ppt);
            HRESULT GetSpacing([Out] out Point ppt);
            HRESULT GetDefaultSpacing(out Point ppt);
            HRESULT GetAutoArrange();
            HRESULT SelectItem(int iItem, uint dwFlags);
            HRESULT SelectAndPositionItems(uint cidl, [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] IntPtr[] apidl, ref Point apt, uint dwFlags);
        }
        
        [ComImport()]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
        public interface IShellItem
        {
            HRESULT BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, ref IntPtr ppv);
            HRESULT GetParent(ref IShellItem ppsi);
            [PreserveSig]
            HRESULT GetDisplayName(SIGDN sigdnName, ref System.Text.StringBuilder ppszName);
            HRESULT GetAttributes(uint sfgaoMask, ref uint psfgaoAttribs);
            HRESULT Compare(IShellItem psi, uint hint, ref int piOrder);
        }

        public enum SIGDN : int
        {
            SIGDN_NORMALDISPLAY = 0x0,
            SIGDN_PARENTRELATIVEPARSING = unchecked((int)0x80018001),
            SIGDN_DESKTOPABSOLUTEPARSING = unchecked((int)0x80028000),
            SIGDN_PARENTRELATIVEEDITING = unchecked((int)0x80031001),
            SIGDN_DESKTOPABSOLUTEEDITING = unchecked((int)0x8004C000),
            SIGDN_FILESYSPATH = unchecked((int)0x80058000),
            SIGDN_URL = unchecked((int)0x80068000),
            SIGDN_PARENTRELATIVEFORADDRESSBAR = unchecked((int)0x8007C001),
            SIGDN_PARENTRELATIVE =  unchecked((int)0x80080001)
        }

        [ComImport()]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("b63ea76d-1f85-456f-a19c-48159efa858b")]
        public interface IShellItemArray
        {
            // Function BindToHandler(pbc As IBindCtx, ByRef bhid As Guid, ByRef riid As Guid, ByRef ppvOut As IntPtr) As HRESULT
            HRESULT BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, ref IntPtr ppvOut);
            HRESULT GetPropertyStore(GETPROPERTYSTOREFLAGS flags, ref Guid riid, ref IntPtr ppv);
            HRESULT GetPropertyDescriptionList(REFPROPERTYKEY keyType, ref Guid riid, ref IntPtr ppv);
            // Function GetAttributes(AttribFlags As SIATTRIBFLAGS, sfgaoMask As SFGAOF, ByRef psfgaoAttribs As SFGAOF) As HRESULT
            HRESULT GetAttributes(SIATTRIBFLAGS AttribFlags, int sfgaoMask, ref int psfgaoAttribs);
            HRESULT GetCount(ref int pdwNumItems);
            HRESULT GetItemAt(int dwIndex, out IShellItem ppsi);
            // Function EnumItems(ByRef ppenumShellItems As IEnumShellItems) As HRESULT
            HRESULT EnumItems(ref IntPtr ppenumShellItems);
        }

        public enum GETPROPERTYSTOREFLAGS
        {
            GPS_DEFAULT = 0,
            GPS_HANDLERPROPERTIESONLY = 0x1,
            GPS_READWRITE = 0x2,
            GPS_TEMPORARY = 0x4,
            GPS_FASTPROPERTIESONLY = 0x8,
            GPS_OPENSLOWITEM = 0x10,
            GPS_DELAYCREATION = 0x20,
            GPS_BESTEFFORT = 0x40,
            GPS_NO_OPLOCK = 0x80,
            GPS_PREFERQUERYPROPERTIES = 0x100,
            GPS_EXTRINSICPROPERTIES = 0x200,
            GPS_EXTRINSICPROPERTIESONLY = 0x400,
            GPS_MASK_VALID = 0x7FF
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct REFPROPERTYKEY
        {
            private Guid fmtid;
            private int pid;
            public Guid FormatId
            {
                get
                {
                    return this.fmtid;
                }
            }
            public int PropertyId
            {
                get
                {
                    return this.pid;
                }
            }
            public REFPROPERTYKEY(Guid formatId, int propertyId)
            {
                this.fmtid = formatId;
                this.pid = propertyId;
            }
            public static readonly REFPROPERTYKEY PKEY_DateCreated = new REFPROPERTYKEY(new Guid("B725F130-47EF-101A-A5F1-02608C9EEBAC"), 15);
        }

        public enum SIATTRIBFLAGS
        {
            SIATTRIBFLAGS_AND = 0x1,
            SIATTRIBFLAGS_OR = 0x2,
            SIATTRIBFLAGS_APPCOMPAT = 0x3,
            SIATTRIBFLAGS_MASK = 0x3,
            SIATTRIBFLAGS_ALLITEMS = 0x4000
        }

        [ComImport()]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("361bbdc7-e6ee-4e13-be58-58e2240c810f")]
        public interface IExplorerBrowserEvents
        {
            [PreserveSig]           
            HRESULT OnNavigationPending(IntPtr pidlFolder);

            [PreserveSig]
            HRESULT OnViewCreated(IShellView psv);
            //HRESULT OnViewCreated([MarshalAs(UnmanagedType.IUnknown)] object psv);

            [PreserveSig]           
            HRESULT OnNavigationComplete(IntPtr pidlFolder);

            [PreserveSig]          
            HRESULT OnNavigationFailed(IntPtr pidlFolder);
        }

        [ComImport]
        [Guid("000214E3-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IShellView : IOleWindow
        {
            #region <IOleWindow>
            new HRESULT GetWindow(out IntPtr phwnd);
            new HRESULT ContextSensitiveHelp(bool fEnterMode);
            #endregion

            HRESULT TranslateAccelerator(MSG pmsg);
            HRESULT EnableModeless(bool fEnable);
            HRESULT UIActivate(uint uState);
            HRESULT Refresh();
            //HRESULT CreateViewWindow(IShellView psvPrevious, FOLDERSETTINGS pfs, IShellBrowser psb, RECT prcView, out IntPtr pIntPtr);
            HRESULT CreateViewWindow(IShellView psvPrevious, FOLDERSETTINGS pfs, IntPtr psb, RECT prcView, out IntPtr pIntPtr);
            HRESULT DestroyViewWindow();
            HRESULT GetCurrentInfo(out FOLDERSETTINGS pfs);
            //HRESULT AddPropertySheetPages(int dwReserved, LPFNSVADDPROPSHEETPAGE pfn, IntPtr lparam);
            HRESULT AddPropertySheetPages(int dwReserved, IntPtr pfn, IntPtr lparam);
            HRESULT SaveViewState();
            HRESULT SelectItem(IntPtr pidlItem, SVSIF uFlags);
            HRESULT GetItemObject(uint uItem, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out object ppv);
        };

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public int wParam;
            public IntPtr lParam;
            public int time;
            public Point pt;
        }

        public enum SVGIO
        {
            SVGIO_BACKGROUND = 0,
            SVGIO_SELECTION = 0x1,
            SVGIO_ALLVIEW = 0x2,
            SVGIO_CHECKED = 0x3,
            SVGIO_TYPE_MASK = 0xf,
            SVGIO_FLAG_VIEWORDER = unchecked((int)0x80000000)
        }

        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleWindow
        {
            HRESULT GetWindow(out IntPtr phwnd);
            HRESULT ContextSensitiveHelp(bool fEnterMode);
        }

        public enum SVSIF
        {
            SVSI_DESELECT = 0,
            SVSI_SELECT = 0x1,
            SVSI_EDIT = 0x3,
            SVSI_DESELECTOTHERS = 0x4,
            SVSI_ENSUREVISIBLE = 0x8,
            SVSI_FOCUSED = 0x10,
            SVSI_TRANSLATEPT = 0x20,
            SVSI_SELECTIONMARK = 0x40,
            SVSI_POSITIONITEM = 0x80,
            SVSI_CHECK = 0x100,
            SVSI_CHECK2 = 0x200,
            SVSI_KEYBOARDSELECT = 0x401,
            SVSI_NOTAKEFOCUS = 0x40000000
        }


        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("User32.dll", SetLastError = true,CharSet = CharSet.Auto)]
        public static extern int GetSystemMetrics(int nIndex);

        public const int SM_CXSCREEN = 0;
        public const int SM_CYSCREEN = 1;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int SetWindowRgn(IntPtr hWnd, IntPtr hRgn, bool bRedraw);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateRoundRectRgn(int x1, int y1, int x2, int y2, int cx, int cy);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CreateRectRgn(int x1, int y1, int x2, int y2);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int CombineRgn(IntPtr hrgnDest, IntPtr hrgnSrc1, IntPtr hrgnSrc2, int iMode);

        [DllImport("Gdi32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int SelectClipRgn(IntPtr hdc, IntPtr hrgn);

        public const int RGN_AND = 1;
        public const int RGN_OR = 2;
        public const int RGN_XOR = 3;
        public const int RGN_DIFF = 4;
        public const int RGN_COPY = 5;
        public const int RGN_MIN = RGN_AND;
        public const int RGN_MAX = RGN_COPY;

        public const int ERROR = 0;
        public const int NULLREGION = 1;
        public const int SIMPLEREGION = 2;
        public const int COMPLEXREGION = 3;

        [DllImport("Gdi32.dll", SetLastError = true)]
        public static extern bool DeleteObject(IntPtr hObject);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        public const int GW_HWNDFIRST = 0;
        public const int GW_HWNDLAST = 1;
        public const int GW_HWNDNEXT = 2;
        public const int GW_HWNDPREV = 3;
        public const int GW_OWNER = 4;
        public const int GW_CHILD = 5;
        public const int GW_ENABLEDPOPUP = 6;

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool EnableWindow(IntPtr hWnd, bool bEnable);  

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, [In, Out] ref RECT pvParam, uint fWinIni);

        public const int SPI_GETWORKAREA = 0x30;

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern HRESULT SHILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, ref IntPtr ppIdl, ref uint rgflnOut);

        [DllImport("Shell32.dll", SetLastError = true)]
        public static extern HRESULT SHGetKnownFolderIDList(ref Guid rfid, int dwFlags, IntPtr hToken, ref IntPtr ppidl);

        [DllImport("Shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool SHGetPathFromIDList(IntPtr pidl, [MarshalAs(UnmanagedType.LPTStr)] System.Text.StringBuilder pszPath);

        [DllImport("Dwmapi.dll", SetLastError = true, ExactSpelling = true)]
        public static extern HRESULT DwmSetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE dwAttribute, ref IntPtr pvAttribute, int cbAttribute);

        public enum DWMWINDOWATTRIBUTE
        {
            DWMWA_NCRENDERING_ENABLED = 1,              // [get] Is non-client rendering enabled/disabled
            DWMWA_NCRENDERING_POLICY,                   // [set] DWMNCRENDERINGPOLICY - Non-client rendering policy
            DWMWA_TRANSITIONS_FORCEDISABLED,            // [set] Potentially enable/forcibly disable transitions
            DWMWA_ALLOW_NCPAINT,                        // [set] Allow contents rendered in the non-client area to be visible on the DWM-drawn frame.
            DWMWA_CAPTION_BUTTON_BOUNDS,                // [get] Bounds of the caption button area in window-relative space.
            DWMWA_NONCLIENT_RTL_LAYOUT,                 // [set] Is non-client content RTL mirrored
            DWMWA_FORCE_ICONIC_REPRESENTATION,          // [set] Force this window to display iconic thumbnails.
            DWMWA_FLIP3D_POLICY,                        // [set] Designates how Flip3D will treat the window.
            DWMWA_EXTENDED_FRAME_BOUNDS,                // [get] Gets the extended frame bounds rectangle in screen space
            DWMWA_HAS_ICONIC_BITMAP,                    // [set] Indicates an available bitmap when there is no better thumbnail representation.
            DWMWA_DISALLOW_PEEK,                        // [set] Don't invoke Peek on the window.
            DWMWA_EXCLUDED_FROM_PEEK,                   // [set] LivePreview exclusion information
            DWMWA_CLOAK,                                // [set] Cloak or uncloak the window
            DWMWA_CLOAKED,                              // [get] Gets the cloaked state of the window
            DWMWA_FREEZE_REPRESENTATION,                // [set] BOOL, Force this window to freeze the thumbnail without live update
            DWMWA_PASSIVE_UPDATE_MODE,                  // [set] BOOL, Updates the window only when desktop composition runs for other reasons
            DWMWA_USE_HOSTBACKDROPBRUSH,                // [set] BOOL, Allows the use of host backdrop brushes for the window.
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20,         // [set] BOOL, Allows a window to either use the accent color, or dark, according to the user Color Mode preferences.
            DWMWA_MICA_EFFECT = 1029,
            DWMWA_LAST
        };

        //[DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        //public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, ref HIGHCONTRASTW pvParam, uint fWinIni);

        [DllImport("UxTheme.dll", SetLastError = true, EntryPoint = "#135", CallingConvention = CallingConvention.Winapi)]
        public static extern bool SetPreferredAppMode(PREFERREDAPPMODE appMode);

        public enum PREFERREDAPPMODE
        {
            Default,
            AllowDark,
            ForceDark,
            ForceLight,
            Max
        }

        [DllImport("UxTheme.dll", SetLastError = true, EntryPoint = "#133", CallingConvention = CallingConvention.Winapi)]
        public static extern bool AllowDarkModeForWindow(IntPtr hWnd, bool allow);

        [DllImport("UxTheme.dll", SetLastError = true, EntryPoint = "#104", CallingConvention = CallingConvention.Winapi)]
        public static extern void RefreshImmersiveColorPolicyState();

        [DllImport("UxTheme.dll", SetLastError = true, EntryPoint = "#137", CallingConvention = CallingConvention.Winapi)]
        public static extern bool IsDarkModeAllowedForWindow(IntPtr hWnd); 

        [DllImport("UxTheme.dll", SetLastError = true, EntryPoint = "#132", CallingConvention = CallingConvention.Winapi)]
        public static extern bool ShouldAppsUseDarkMode();

        [DllImport("UxTheme.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int SetWindowTheme(IntPtr hwnd, string pszSubAppName, string pszSubIdList);

 


        IntPtr hWnd = IntPtr.Zero;
        private Microsoft.UI.Windowing.AppWindow _apw;
        private IntPtr hWndChild = IntPtr.Zero;

        CExplorerBrowserHost ebh;
        private int nXControl = 10, nYControl = 10;
        private int nWidthControl = 1000, nHeightControl = 460;
        private IntPtr hWndBrowserControl = IntPtr.Zero;

        public MainWindow()
        {
            this.InitializeComponent();

            SetPreferredAppMode(PREFERREDAPPMODE.AllowDark);            

            hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            Microsoft.UI.WindowId myWndId = Microsoft.UI.Win32Interop.GetWindowIdFromWindow(hWnd);
            _apw = Microsoft.UI.Windowing.AppWindow.GetFromWindowId(myWndId);

            hWndChild = FindWindowEx(hWnd, IntPtr.Zero, "Microsoft.UI.Content.ContentWindowSiteBridge", null);

            IntPtr pAttribute = (IntPtr)1;
            HRESULT hr = DwmSetWindowAttribute(hWnd, DWMWINDOWATTRIBUTE.DWMWA_USE_IMMERSIVE_DARK_MODE, ref pAttribute, Marshal.SizeOf(typeof(IntPtr)));

            RECT rc = new RECT(nXControl, nYControl, nWidthControl + nXControl, nHeightControl + nXControl);
            ebh = new CExplorerBrowserHost(this, hWnd, rc);
            hWndBrowserControl = FindWindowEx(hWnd, IntPtr.Zero, "ExplorerBrowserControl", null);

            //this.SizeChanged += MainWindow_SizeChanged;
            //this.VisibilityChanged += MainWindow_VisibilityChanged;
          
            //RefreshImmersiveColorPolicyState();
            bool bRet = AllowDarkModeForWindow(hWndBrowserControl, true);
            //SetWindowTheme(hWndBrowserControl, "DarkMode_Explorer", null);
            SetWindowTheme(hWndBrowserControl, "Explorer", null);
            SetRegion(true);

            _apw.Resize(new Windows.Graphics.SizeInt32(1038, 640));
            //_apw.Move(new Windows.Graphics.PointInt32(600, 400));
            CenterToScreen(hWnd);
        }

        //private void MainWindow_VisibilityChanged(object sender, WindowVisibilityChangedEventArgs args)
        //{
        //    SetRegion(true);
        //}

        //private void MainWindow_SizeChanged(object sender, WindowSizeChangedEventArgs args)
        //{
        //    SetRegion(true);
        //}

        private void myButton_Click(object sender, RoutedEventArgs e)
        {         
            //myButton.Content = "Clicked";
            Click();
        }

        private async void Click()
        {
            SetRegion(false);
            EnableWindow(hWndBrowserControl, false);
            StackPanel sp = new StackPanel();
            // https://www.unicode.org/emoji/charts/full-emoji-list.html
            FontIcon fi = new FontIcon()
            {
                FontFamily = new FontFamily("Segoe UI Emoji"),
                Glyph = "\U0001F439",
                //Glyph = "\U00002699", // Gear           
                FontSize = 50
            };
            sp.Children.Add(fi);
            TextBlock tb = new TextBlock();
            tb.HorizontalAlignment = HorizontalAlignment.Center;
            tb.Text = "You clicked on the Button !";
            sp.Children.Add(tb);
            ContentDialog cd = new ContentDialog()
            {
                Title = "Information",
                Content = sp,
                CloseButtonText = "Ok"
            };
            cd.XamlRoot = this.Content.XamlRoot;
            var res = await cd.ShowAsync();
            EnableWindow(hWndBrowserControl, true);
            SetRegion(true);
        }

        private void OnNavigationComplete(IntPtr pidlFolder)
        {
            //System.Text.StringBuilder sbFolderName = new System.Text.StringBuilder(260);
            //SHGetPathFromIDList(pidlFolder, sbFolderName);
            //tbItem.Text = sbFolderName.ToString();

            if (ebh != null)
            {
                IntPtr pFolderViewPtr = IntPtr.Zero;
                Guid IID_IFolderView = typeof(IFolderView).GUID;
                HRESULT hr = ebh.pExplorerBrowser.GetCurrentView(IID_IFolderView, ref pFolderViewPtr);
                if (hr == HRESULT.S_OK)
                {
                    IFolderView pFolderView = (IFolderView)Marshal.GetObjectForIUnknown(pFolderViewPtr);
                    object oShellItem = null;
                    Guid IID_IShellItem = typeof(IShellItem).GUID;
                    hr = pFolderView.GetFolder(ref IID_IShellItem, out oShellItem);
                    if (hr == HRESULT.S_OK)
                    {
                        IShellItem pShellItem = (IShellItem)oShellItem;
                        System.Text.StringBuilder sbFolderName = new System.Text.StringBuilder(260);
                        // hr = pShellItem.GetDisplayName(SIGDN.SIGDN_PARENTRELATIVEFORADDRESSBAR, sbItemName);  
                        hr = pShellItem.GetDisplayName(SIGDN.SIGDN_DESKTOPABSOLUTEPARSING, ref sbFolderName);
                        tbItem.Text = sbFolderName.ToString();                       
                    }
                    Marshal.ReleaseComObject(pFolderView);
                }
            }
        }

        private void buttonCurrentItem_Click(object sender, RoutedEventArgs e)
        {
            IntPtr pFolderViewPtr = IntPtr.Zero;
            Guid IID_IFolderView = typeof(IFolderView).GUID;
            HRESULT hr = ebh.pExplorerBrowser.GetCurrentView(IID_IFolderView, ref pFolderViewPtr);
            if (hr == HRESULT.S_OK)
            {
                IFolderView pFolderView = (IFolderView)Marshal.GetObjectForIUnknown(pFolderViewPtr);
                object oShellItem = null;
                Guid IID_IShellItem = typeof(IShellItem).GUID;
                hr = pFolderView.GetFolder(ref IID_IShellItem, out oShellItem);
                if (hr == HRESULT.S_OK)
                { 
                    int nItem = 0;
                    hr = pFolderView.GetFocusedItem(out nItem);
 
                    object oShellItemArray = null;
                    Guid IID_IShellItemArray = typeof(IShellItemArray).GUID;
                    textBlockItems.Text = "";
                    // 	hr	0x80070490	 ERROR_NOT_FOUND
                    hr = pFolderView.Items((uint)SVGIO.SVGIO_SELECTION, ref IID_IShellItemArray, out oShellItemArray);
                    if (hr == HRESULT.S_OK)
                    {
                        IShellItemArray pShellItemArray = (IShellItemArray)oShellItemArray;
                        int nNbItems = 0;
                        pShellItemArray.GetCount(ref nNbItems);
                        for (int i = 0; i < nNbItems; i++)
                        {
                            IShellItem psi;
                            pShellItemArray.GetItemAt(i, out psi);
                            System.Text.StringBuilder sbItemName = new System.Text.StringBuilder(260);
                            psi.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY, ref sbItemName);
                            textBlockItems.Text += String.Format("[{0}] ", sbItemName.ToString());                       
                        }
                        Marshal.ReleaseComObject(pShellItemArray);
                    }                   
                }
                Marshal.ReleaseComObject(pFolderView);
            }
        }

        private void SetRegion(bool bRegion)
        {
            if (bRegion)
            {
                RECT rc;
                GetClientRect(hWnd, out rc);              
                int nScreenWidth = GetSystemMetrics(SM_CXSCREEN);
                int nScreenHeight = GetSystemMetrics(SM_CYSCREEN);
                IntPtr WindowRgn = CreateRectRgn(0, 0, nScreenWidth, nScreenHeight);
                IntPtr HoleRgn = CreateRectRgn(nXControl, nYControl, nXControl + nWidthControl, nYControl + nHeightControl);
                CombineRgn(WindowRgn, WindowRgn, HoleRgn, RGN_DIFF);
                SetWindowRgn(hWndChild, WindowRgn, true);
                DeleteObject(HoleRgn);
            }
            else
                SetWindowRgn(hWndChild, IntPtr.Zero, true);
        }

        private void CenterToScreen(IntPtr hWnd)
        {
            RECT rcWorkArea = new RECT();
            SystemParametersInfo(SPI_GETWORKAREA, 0, ref rcWorkArea, 0);
            RECT rc;
            GetWindowRect(hWnd, out rc);
            int nX = System.Convert.ToInt32((rcWorkArea.left + rcWorkArea.right) / (double)2 - (rc.right - rc.left) / (double)2);
            int nY = System.Convert.ToInt32((rcWorkArea.top + rcWorkArea.bottom) / (double)2 - (rc.bottom - rc.top) / (double)2);
            //SetWindowPos(hWnd, IntPtr.Zero, nX, nY, -1, -1, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
            SetWindowPos(hWnd, IntPtr.Zero, nX, nY, -1, -1, SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);            
        }
    }
}
