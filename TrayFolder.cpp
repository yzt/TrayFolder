#include <Windows.h>
#include <Windowsx.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <PathCch.h>
#include <CommCtrl.h>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "pathcch.lib")
#pragma comment(lib, "shlwapi.lib")

// #define RANDOMLY_GENERATE_GUID
#if defined(RANDOMLY_GENERATE_GUID)
    #include <bcrypt.h>
    #pragma comment(lib, "bcrypt.lib")
#endif

// #define ALSO_LOG_TO_FILE
// #define ALSO_LOG_TO_CONSOLE
#if defined(ALSO_LOG_TO_CONSOLE) || defined(ALSO_LOG_TO_FILE)
    #include <cstdio>
#endif

#define APP_NAME L"TrayFolder"

#define T_DO(x_) L##x_
#define _T(x_)   T_DO(x_)

#define CHECK_WINAPI(expr_) \
    ReportError((expr_), (decltype(expr_))0, true, APP_NAME, L"", _T(#expr_), _T(__FUNCTION__), _T(__FILE__), __LINE__)

#define LOG_INFO(fmt_, ...)  Log("(I) " fmt_ "\n" __VA_OPT__(, ) __VA_ARGS__)
#define LOG_DEBUG(fmt_, ...) Log("(D) " fmt_ "\n" __VA_OPT__(, ) __VA_ARGS__)

// TODO(yzt): Add DEFER

constexpr UINT NotifyMessage = WM_USER + 0x44;
#if defined(RANDOMLY_GENERATE_GUID)
// A randomly-generated GUID will cause Windows to forget the user's choice about always showing the app icon in the
// tray
GUID TrayIconGuid = {};
#else
// {B89F179A-8BAB-4140-BFFE-A79D08AFED29}
constexpr GUID kTrayIconGuid = {
    0xb89f179a,
    0x8bab,
    0x4140,
    {0xbf, 0xfe, 0xa7, 0x9d, 0x8, 0xaf, 0xed, 0x29}
};
#endif

struct FolderData {
    wchar_t dir[MAX_PATH];
    int     count;
    struct {
        wchar_t file[MAX_PATH];
        wchar_t title[243];
        bool    isDir;
        HICON   icon;
        HBITMAP bitmap;
    } entries[256];
};

// Global data:
static FolderData    g_data;
static wchar_t*      g_path;
static SHSTOCKICONID g_icon     = SHSTOCKICONID::SIID_FOLDER;
static GUID          g_guid     = kTrayIconGuid;
static bool          g_poppedUp = false;

static void Log(char const* fmt, ...) {
#if defined(ALSO_LOG_TO_FILE)
    static FILE* s_logf = [] {
        FILE* ret = nullptr;
        ::fopen_s(&ret, APP_NAME ".log", "ab");
        return ret;
    }();
#endif
    char    buff[512];
    va_list args;
    va_start(args, fmt);
    ::wvnsprintfA(buff, _countof(buff), fmt, args);
    buff[_countof(buff) - 1] = '\0';
    va_end(args);
    ::OutputDebugStringA(buff);

#if defined(ALSO_LOG_TO_FILE)
    ::fputs(buff, s_logf);
    ::fflush(s_logf);
#endif
#if defined(ALSO_LOG_TO_CONSOLE)
    ::fputs(buff, stderr);
    ::fflush(stderr);
#endif
}

template<typename T>
static T ReportError(
    T const&       value,
    T const&       errorVal,
    bool           isFatal,
    wchar_t const* header,
    wchar_t const* message,
    wchar_t const* expr,
    wchar_t const* func,
    wchar_t const* /*file*/,
    int line
) {
    bool const cond = (value != errorVal);
    if (!cond) {
        //__debugbreak();
        wchar_t mstr[1024];
        wchar_t estr[512];

        DWORD const ec = ::GetLastError();
        ::FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | 0, nullptr, ec, 0, estr, _countof(estr), nullptr);
        estr[_countof(estr) - 1] = '\0';
        ::wsprintfW(
            mstr, L"%s%s failed (in %s@%d):\n\n[0x%08X] %s%s", message, expr, func, line, ec, estr,
            isFatal ? L"\n\nThe application will close now." : L""
        );
        mstr[_countof(mstr) - 1] = '\0';

        ::MessageBoxW(nullptr, mstr, header, MB_OK | MB_ICONERROR);
        if (isFatal) {
            ::ExitProcess(1);
            __assume(0);
        }
    }

    return value;
}

static double NowMs() {
    static double const ToMsFactor = [] {
        LARGE_INTEGER freq;
        ::QueryPerformanceFrequency(&freq);
        return 1000.0 / freq.QuadPart;
    }();

    LARGE_INTEGER counter;
    ::QueryPerformanceCounter(&counter);
    return ToMsFactor * counter.QuadPart;
}

static size_t StrCopy(wchar_t* dst, size_t dstCap, wchar_t const* src, size_t srcCap = ~(size_t)0) {
    size_t i = 0;
    if (dst && dstCap) {
        if (src) {
            for (i = 0; i + 1 < dstCap && i < srcCap && src[i]; ++i) {
                dst[i] = src[i];
            }
        }
        dst[i] = '\0';
    }
    return i;
}

static int StrCompare(wchar_t const* s1, wchar_t const* s2) {
    for (int i = 0; s1[i] || s2[i]; ++i) {
        // This way of case-insensitive comparison is not exactly correct, but who cares
        int const c1 = ('A' <= s1[i] && s1[i] <= 'Z' ? s1[i] -'A' + 'a' : s1[i]);
        int const c2 = ('A' <= s2[i] && s2[i] <= 'Z' ? s2[i] -'A' + 'a' : s2[i]);
        if (c1 != c2)
            return c1 - c2;
    }
    return 0;
}

static bool LoadFolder(FolderData& out, wchar_t const* dir) {
    for (int i = 0; i < out.count; ++i) {
        //::DestroyIcon(out.entries[i].icon);
        CHECK_WINAPI(::DeleteObject(out.entries[i].bitmap));
    }
    out.count = 0;

    StrCopy(out.dir, _countof(out.dir), dir);
    ::PathCchRemoveBackslash(out.dir, _countof(out.dir));

    wchar_t glob[MAX_PATH + 10];
    ::PathCchCombine(glob, _countof(glob), out.dir, L"*");
    WIN32_FIND_DATAW findData   = {};
    HANDLE           findHandle = ::FindFirstFileExW(
        glob, FindExInfoBasic, &findData, FindExSearchNameMatch, nullptr, FIND_FIRST_EX_LARGE_FETCH
    );
    if (findHandle != INVALID_HANDLE_VALUE) {
        do {
            if (findData.cFileName[0] == '.')
                continue;
            // if (findData.nFileSizeLow == 0 && findData.nFileSizeHigh == 0) continue;
            auto& entry = out.entries[out.count++];
            StrCopy(entry.file, _countof(entry.file), findData.cFileName, _countof(findData.cFileName));
            entry.isDir = (0 != (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY));

            // Remove extension for display
            StrCopy(entry.title, _countof(entry.title), findData.cFileName, _countof(findData.cFileName));
            ::PathCchRemoveExtension(entry.title, _countof(entry.title));

            // Make a full path
            wchar_t fullpath[2 * MAX_PATH];
            ::PathCchCombineEx(
                fullpath, _countof(fullpath), out.dir, findData.cFileName,
                PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS
            );

            // Get the icon for the file, and render it to a new bitmap (this produces correct masking, and also removes
            // possible overlays, e.g. link arrow, etc.)
            SHFILEINFOW shinfo       = {};
            DWORD_PTR   sysImageList = CHECK_WINAPI(::SHGetFileInfoW(
                fullpath, findData.dwFileAttributes, &shinfo, sizeof(shinfo),
                SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES
            ));
            IMAGEINFO   imageInfo;
            CHECK_WINAPI(::ImageList_GetImageInfo((HIMAGELIST)sysImageList, shinfo.iIcon, &imageInfo));
            LONG const w          = imageInfo.rcImage.right - imageInfo.rcImage.left;
            LONG const h          = imageInfo.rcImage.bottom - imageInfo.rcImage.top;
            HDC        hDc        = CHECK_WINAPI(::GetDC(nullptr));
            HDC        hMemDc     = CHECK_WINAPI(::CreateCompatibleDC(hDc));
            HBITMAP    hTargetBmp = CHECK_WINAPI(::CreateCompatibleBitmap(hDc, w, h));
            HGDIOBJ    hOrigBmp   = ::SelectObject(hMemDc, hTargetBmp);
            CHECK_WINAPI(ImageList_Draw((HIMAGELIST)sysImageList, shinfo.iIcon, hMemDc, 0, 0, ILD_NORMAL));
            ::SelectObject(hMemDc, hOrigBmp);
            CHECK_WINAPI(::DeleteDC(hMemDc));
            ::ReleaseDC(nullptr, hDc);
            entry.bitmap = hTargetBmp;
        } while (out.count < _countof(out.entries) && ::FindNextFileW(findHandle, &findData));
        ::FindClose(findHandle);
    }
    return true;
}

static void SortFolderData(FolderData& data) {
    int const n = data.count;
    auto      a = data.entries;
    char      temp[sizeof(a[0])];
    for (int i = 0; i < n - 1; ++i) {
        int min = i;
        for (int j = i + 1; j < n; ++j) {
            if ((a[j].isDir == a[min].isDir && StrCompare(a[j].title, a[min].title) < 0)
                || (a[j].isDir != a[min].isDir && a[j].isDir)) {
                min = j;
            }
        }
        if (min != i) {
            ::CopyMemory(temp, &a[min], sizeof(temp));
            ::MoveMemory(&a[i + 1], &a[i], sizeof(temp) * (min - i));
            ::CopyMemory(&a[i], temp, sizeof(temp));
        }
    }
}

static HMENU BuildFolderMenu(FolderData& data) {
    HMENU menu = CHECK_WINAPI(::CreatePopupMenu());
    for (int i = 0; i < data.count; ++i) {
        MENUITEMINFOW item = {
            .cbSize     = sizeof(MENUITEMINFOW),
            .fMask      = 0 | MIIM_BITMAP /*| MIIM_DATA*/ | MIIM_ID /*| MIIM_STATE*/ | MIIM_STRING,
            .wID        = 100 + static_cast<unsigned>(i),
            .dwTypeData = data.entries[i].title,
            //.cch = 1,
            .hbmpItem   = data.entries[i].bitmap,
        };
        CHECK_WINAPI(::InsertMenuItemW(menu, 0xFFFF'FFFF, true, &item));
    }

    MENUITEMINFOW separatorItem = {.cbSize = sizeof(MENUITEMINFOW), .fMask = MIIM_FTYPE, .fType = MFT_SEPARATOR};
    CHECK_WINAPI(::InsertMenuItemW(menu, 0xFFFF'FFFF, true, &separatorItem));

    // wchar_t refreshText[] = L"&Refresh";
    // MENUITEMINFOW refreshItem = {.cbSize = sizeof(MENUITEMINFOW), .fMask = MIIM_ID | MIIM_STRING, .wID = 2,
    // .dwTypeData = refreshText}; CHECK_WINAPI(::InsertMenuItemW(menu, 0xFFFF'FFFF, true, &refreshItem));

    wchar_t       hideText[] = L"&Dismiss";
    MENUITEMINFOW hideItem   = {
          .cbSize = sizeof(MENUITEMINFOW), .fMask = MIIM_ID | MIIM_STRING, .wID = 3, .dwTypeData = hideText
    };
    CHECK_WINAPI(::InsertMenuItemW(menu, 0xFFFF'FFFF, true, &hideItem));

    wchar_t       openText[] = L"&Open Folder";
    MENUITEMINFOW openItem   = {
          .cbSize = sizeof(MENUITEMINFOW), .fMask = MIIM_ID | MIIM_STRING, .wID = 2, .dwTypeData = openText
    };
    CHECK_WINAPI(::InsertMenuItemW(menu, 0xFFFF'FFFF, true, &openItem));

    wchar_t       quitText[] = L"&Quit";
    MENUITEMINFOW quitItem   = {
          .cbSize = sizeof(MENUITEMINFOW), .fMask = MIIM_ID | MIIM_STRING, .wID = 1, .dwTypeData = quitText
    };
    CHECK_WINAPI(::InsertMenuItemW(menu, 0xFFFF'FFFF, true, &quitItem));

    return menu;
}

static int ShowFolderMenu(HWND hWnd, GUID const& iconGuid, HMENU menu) {
    // Get the rectangle for the icon
    RECT                 iconRect = {};
    NOTIFYICONIDENTIFIER niid     = {.cbSize = sizeof(NOTIFYICONIDENTIFIER), .hWnd = hWnd, .guidItem = iconGuid};
    CHECK_WINAPI(SUCCEEDED(::Shell_NotifyIconGetRect(&niid, &iconRect)));

    // Must do this (even though hWnd is not a real window), so the pop-up is dismiss-able
    // FIXME: when this fails (e.g. when start menu is open) you have to use "Dismiss" item in the menu to dismiss it
    // NOTE: (to future me) this does *not* interact with ::GetLastError()
    ::SetForegroundWindow(hWnd);

    TPMPARAMS tpmParams = {.cbSize = sizeof(TPMPARAMS), .rcExclude = iconRect};
    UINT      flags     = /*TPM_LEFTALIGN | TPM_TOPALIGN |*/ TPM_NONOTIFY
               | TPM_RETURNCMD /*| TPM_LEFTBUTTON | TPM_RIGHTBUTTON*/ /*| TPM_NOANIMATION*/ /*| TPM_VERTICAL*/;
    int choice = ::TrackPopupMenuEx(menu, flags, iconRect.left, iconRect.top, hWnd, &tpmParams);

    NOTIFYICONDATAW nidSetVer = {
        .cbSize   = sizeof(NOTIFYICONDATAW),
        .hWnd     = hWnd,
        .uFlags   = NIF_GUID,
        .uVersion = NOTIFYICON_VERSION_4,
        .guidItem = g_guid,
    };
    CHECK_WINAPI(::Shell_NotifyIconW(NIM_SETFOCUS, &nidSetVer));

    return choice;
}

static LRESULT CALLBACK MessageWindowProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
// #define LOG_EVERY_WINDOW_MESSAGE
#if defined(LOG_EVERY_WINDOW_MESSAGE)
    if (NotifyMessage == msg) {
        LOG_DEBUG(
            "Message: hwnd=%08p, event=%04hX:%-20s, id=%hu, x=%hd, y=%hd", hWnd, LOWORD(lParam),
            FromMsg(LOWORD(lParam)), HIWORD(lParam), GET_X_LPARAM(wParam), GET_Y_LPARAM(wParam)
        );
    } else {
        LOG_DEBUG(
            "Message: hwnd=%08p, msg=%08X:%-20s, wparam=%016llX, lparam=%016llX", hWnd, msg, FromMsg(msg), wParam, lParam
        );
    }
#endif

    if (msg == NotifyMessage) {
        switch (LOWORD(lParam)) {
        case NIN_SELECT:       // Left-click
        case WM_CONTEXTMENU: { // Right-click
            if (!g_poppedUp) {
                g_poppedUp      = true;
                double const t0 = NowMs();
                CHECK_WINAPI(LoadFolder(g_data, g_path));
                SortFolderData(g_data);
                double const t1   = NowMs();
                HMENU const  menu = CHECK_WINAPI(BuildFolderMenu(g_data));
                double const t2   = NowMs();
                LOG_DEBUG(
                    "Building menu took %d ms (%d+%d)", int(t2 - t0 + 0.5), int(t1 - t0 + 0.5), int(t2 - t1 + 0.5)
                );
                int const choice = ShowFolderMenu(hWnd, g_guid, menu);
                LOG_INFO("MenuChoice: %d", choice);
                if (choice == 1) {
                    PostQuitMessage(0);
                } else if (choice == 2) {
                    [[maybe_unused]] auto res = (INT_PTR
                    )::ShellExecuteW(nullptr, nullptr, g_data.dir, nullptr, nullptr, SW_SHOWNORMAL);
                } else if (choice - 100 >= 0 && choice - 100 < g_data.count) {
                    auto const& entry = g_data.entries[choice - 100];
                    wchar_t     path[MAX_PATH * 2];
                    ::PathCchCombineEx(
                        path, _countof(path), g_data.dir, entry.file,
                        PATHCCH_ALLOW_LONG_PATHS | PATHCCH_FORCE_ENABLE_LONG_NAME_PROCESS
                    );
                    LOG_DEBUG("Executing: %S", path);
                    [[maybe_unused]] auto res = (INT_PTR
                    )::ShellExecuteW(nullptr, nullptr, path, nullptr, nullptr, SW_SHOWNORMAL);
                    LOG_DEBUG("%zd", res);
                }
                g_poppedUp = false;
            }
            return ::DefWindowProcW(hWnd, msg, wParam, lParam);
        } break;
        }
    }
    return ::DefWindowProcW(hWnd, msg, wParam, lParam);
}

int APIENTRY
WinMain(_In_ HINSTANCE /*hInst*/, _In_opt_ HINSTANCE /*hPrevInst*/, _In_ LPSTR /*lpCmdLine*/, _In_ int /*nCmdShow*/) {
    constexpr wchar_t ClassName[]  = APP_NAME L"Class";
    constexpr wchar_t WindowName[] = APP_NAME L"Window";
    constexpr UINT    TrayIconId   = 42;
#if defined(RANDOMLY_GENERATE_GUID)
    CHECK_WINAPI(SUCCEEDED(
        BCryptGenRandom(nullptr, (PUCHAR)&TrayIconGuid, sizeof(TrayIconGuid), BCRYPT_USE_SYSTEM_PREFERRED_RNG)
    ));
#endif

    // Process command line
    int       argc = 0;
    wchar_t** argv = ::CommandLineToArgvW(::GetCommandLineW(), &argc);
    if (argc <= 1) {
        ::MessageBoxW(
            nullptr,
            L"Usage:\n  " APP_NAME " <folderPath> [iconNumber]\n\nwhere iconNumber is any of SHSTOCKICONID *integer* "
            "values (defaults to 3, which is an image of a folder)",
            APP_NAME, MB_OK | MB_ICONINFORMATION
        );
        ::ExitProcess(0);
    }
    g_path = argv[1];
    if (argc >= 2) {
        g_icon = SHSTOCKICONID(::StrToIntW(argv[2]));
    }

    // Constant in only available in Win10 1809(?) or later
    CHECK_WINAPI(::SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2));

    // Init COM for ShellExecute
    CHECK_WINAPI(SUCCEEDED(::CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)));

    HMODULE instance = CHECK_WINAPI(::GetModuleHandleW(nullptr));
    LOG_INFO("Instance: %016p", instance);

    WNDCLASSEXW windowClass = {
        .cbSize        = sizeof(WNDCLASSEXW),
        .lpfnWndProc   = MessageWindowProc,
        .hInstance     = instance,
        .lpszClassName = ClassName,
    };
    CHECK_WINAPI(::RegisterClassExW(&windowClass));

    HWND window = CHECK_WINAPI(
        ::CreateWindowExW(0, ClassName, WindowName, 0, 0, 0, 0, 0, HWND_MESSAGE, nullptr, instance, nullptr)
    );
    LOG_INFO("Window: %016p", window);
    CHECK_WINAPI(::UpdateWindow(window));

    SHSTOCKICONINFO stockIconInfo = {.cbSize = sizeof(SHSTOCKICONINFO)};
    if (!SUCCEEDED(::SHGetStockIconInfo(g_icon, SHGSI_ICON | SHGSI_LARGEICON, &stockIconInfo))) {
        CHECK_WINAPI(SUCCEEDED(::SHGetStockIconInfo(SIID_FOLDEROPEN, SHGSI_ICON | SHGSI_LARGEICON, &stockIconInfo)));
    }
    HICON hTrayIcon = CHECK_WINAPI(HICON(stockIconInfo.hIcon));

    // stockIconInfo = {.cbSize = sizeof(SHSTOCKICONINFO)};
    // CHECK_WINAPI(SUCCEEDED(::SHGetStockIconInfo(SIID_DRIVE35, SHGSI_ICON | SHGSI_LARGEICON, &stockIconInfo)));
    // HICON hIconForBalloon = CHECK_WINAPI(HICON(stockIconInfo.hIcon));

    // Add the tray area icon...
    NOTIFYICONDATAW nidAdd = {
        .cbSize           = sizeof(NOTIFYICONDATAW),
        .hWnd             = window,
        .uID              = TrayIconId,
        .uFlags           = NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_STATE /*| NIF_INFO*/ | NIF_GUID | NIF_SHOWTIP,
        .uCallbackMessage = NotifyMessage,
        .hIcon            = hTrayIcon,
        .szTip            = L"Just a tip...",
        .dwState          = 0 /*| NIS_SHAREDICON*/,
        .dwStateMask      = 0,
        .guidItem         = g_guid,
    };
    StrCopy(nidAdd.szTip, _countof(nidAdd.szTip), g_path);
    BOOL addShellNotifyIconResult;
    for (int i = 0; i < 256; ++i) {
        if (FALSE != (addShellNotifyIconResult = ::Shell_NotifyIconW(NIM_ADD, &nidAdd)))
            break;
        // The failure could be due to GUID clash (running another instance of this program)
        g_guid.Data4[7]++;
        nidAdd.guidItem = g_guid;
    }
    CHECK_WINAPI(addShellNotifyIconResult);

    // Request "modern" (Win2K+) behavior...
    NOTIFYICONDATAW nidSetVer = {
        .cbSize   = sizeof(NOTIFYICONDATAW),
        .hWnd     = window,
        .uID      = TrayIconId,
        .uFlags   = NIF_GUID,
        .uVersion = NOTIFYICON_VERSION_4,
        .guidItem = g_guid,
    };
    CHECK_WINAPI(::Shell_NotifyIconW(NIM_SETVERSION, &nidSetVer));

    // Announce app start and direct attention to the icon
    NOTIFYICONDATAW nidPopUp = {
        .cbSize       = sizeof(NOTIFYICONDATAW),
        .hWnd         = window,
        .uID          = TrayIconId,
        .uFlags       = NIF_INFO | NIF_GUID | NIF_SHOWTIP,
        .szInfo       = L"Ballooning the notification; should have been overwritten",
        .uTimeout     = 3'000,
        .szInfoTitle  = APP_NAME L" is running",
        .dwInfoFlags  = NIIF_USER | NIIF_LARGE_ICON | NIIF_NOSOUND | NIIF_RESPECT_QUIET_TIME,
        .guidItem     = g_guid,
        .hBalloonIcon = nullptr, // hIconForBalloon,
    };
    StrCopy(nidPopUp.szInfo, _countof(nidPopUp.szInfo), g_path);
    CHECK_WINAPI(::Shell_NotifyIconW(NIM_MODIFY, &nidPopUp));

    for (;;) {
        MSG        msg;
        BOOL const gm_ret = ::GetMessageW(&msg, nullptr, 0, 0);
        if (0 == gm_ret || -1 == gm_ret) {
            break;
        }
        ::TranslateMessage(&msg);
        ::DispatchMessageW(&msg);
    }

    for (int i = 0; i < g_data.count; ++i) {
        //::DestroyIcon(g_data.entries[i].icon);
        CHECK_WINAPI(::DeleteObject(g_data.entries[i].bitmap));
    }
    CHECK_WINAPI(::DestroyWindow(window));
    CHECK_WINAPI(::UnregisterClassW(ClassName, instance));
    // CHECK_WINAPI(::DestroyIcon(hIconForBalloon));
    CHECK_WINAPI(::DestroyIcon(hTrayIcon));
    ::CoUninitialize();
    ::LocalFree(argv);
    return 0;
}

#if 0 && defined(_CONSOLE)
// Used to be run as a console app, for ease of debugging
int main() {
    return WinMain(::GetModuleHandleW(nullptr), nullptr, ::GetCommandLineA(), 0);
}
#endif
