/*
 * Outlook to Tray - Main Application
 * Manages system tray icon and coordinates with hook DLL
 */

#define _WIN32_WINNT 0x0600
#define NTDDI_VERSION 0x06000000

#include <windows.h>
#include <shellapi.h>
#include <tlhelp32.h>
#include <tchar.h>
#include <string>
#include <thread>
#include <psapi.h>
#include "resource.h"

#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "Psapi.lib")

// Menu item IDs
#define ID_TRAY_ICON        1001
#define ID_TRAY_RESTORE     1002
#define ID_TRAY_AUTOSTART   1003
#define ID_TRAY_ABOUT       1004
#define ID_TRAY_EXIT        1005
#define WM_TRAYICON         (WM_USER + 1)

// DLL function types
typedef BOOL (*InstallHookProc)(HINSTANCE);
typedef BOOL (*UninstallHookProc)();
typedef HWND (*GetHiddenOutlookWindowProc)();
typedef void (*ClearHiddenWindowProc)();

// Globals
HINSTANCE g_hInstance = NULL;
HWND g_hwnd = NULL;
NOTIFYICONDATA g_nid = {};
HMENU g_hMenu = NULL;
HMODULE g_hDll = NULL;
HICON g_hIcon = NULL;
bool g_running = true;

// DLL function pointers
InstallHookProc g_InstallHook = NULL;
UninstallHookProc g_UninstallHook = NULL;
GetHiddenOutlookWindowProc g_GetHiddenOutlookWindow = NULL;
ClearHiddenWindowProc g_ClearHiddenWindow = NULL;

// Debug helper
void DebugMsg(const wchar_t* msg) {
    OutputDebugString(msg);
    OutputDebugString(L"\n");
}

// Find process ID by name
DWORD FindProcessId(const std::wstring& processName) {
    DWORD processId = 0;
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap != INVALID_HANDLE_VALUE) {
        PROCESSENTRY32 pe32 = {};
        pe32.dwSize = sizeof(PROCESSENTRY32);
        if (Process32First(hSnap, &pe32)) {
            do {
                if (_wcsicmp(pe32.szExeFile, processName.c_str()) == 0) {
                    processId = pe32.th32ProcessID;
                    break;
                }
            } while (Process32Next(hSnap, &pe32));
        }
        CloseHandle(hSnap);
    }
    return processId;
}

// Check if autostart is enabled
bool IsAutoStartEnabled() {
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
        0, KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS) {
        DWORD type, size = 0;
        LONG result = RegQueryValueEx(hKey, L"OutlookToTray", NULL, &type, NULL, &size);
        RegCloseKey(hKey);
        return result == ERROR_SUCCESS;
    }
    return false;
}

// Toggle autostart in registry
void ToggleAutoStart() {
    HKEY hKey;
    if (RegOpenKeyEx(HKEY_CURRENT_USER,
        L"Software\\Microsoft\\Windows\\CurrentVersion\\Run",
        0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {

        if (IsAutoStartEnabled()) {
            RegDeleteValue(hKey, L"OutlookToTray");
            MessageBox(g_hwnd, L"Removed from startup.", L"Outlook to Tray", MB_OK | MB_ICONINFORMATION);
        }
        else {
            wchar_t exePath[MAX_PATH];
            GetModuleFileName(NULL, exePath, MAX_PATH);
            RegSetValueEx(hKey, L"OutlookToTray", 0, REG_SZ,
                (BYTE*)exePath, (DWORD)(wcslen(exePath) + 1) * sizeof(wchar_t));
            MessageBox(g_hwnd, L"Added to startup.", L"Outlook to Tray", MB_OK | MB_ICONINFORMATION);
        }
        RegCloseKey(hKey);
    }
}

// Restore hidden Outlook window
void RestoreOutlookWindow() {
    DebugMsg(L"RestoreOutlookWindow called");
    if (g_GetHiddenOutlookWindow) {
        HWND hwnd = g_GetHiddenOutlookWindow();
        DebugMsg(hwnd ? L"Got hidden window handle" : L"No hidden window");
        if (hwnd && IsWindow(hwnd)) {
            ShowWindow(hwnd, SW_SHOW);
            ShowWindow(hwnd, SW_RESTORE);
            SetForegroundWindow(hwnd);
            BringWindowToTop(hwnd);
            if (g_ClearHiddenWindow) {
                g_ClearHiddenWindow();
            }
            DebugMsg(L"Window restored");
        }
        else {
            // No hidden window - check if Outlook is running
            DWORD pid = FindProcessId(L"olk.exe");
            if (pid) {
                // Outlook is running but no hidden window tracked
                MessageBox(NULL, L"Outlook is running but window not found.\nPlease open Outlook manually.",
                          L"Outlook to Tray", MB_OK | MB_ICONINFORMATION);
            } else {
                // Outlook not running - launch it
                DebugMsg(L"Launching Outlook");
                // Use shell to open the new Outlook app
                ShellExecute(NULL, L"open", L"ms-outlook:", NULL, NULL, SW_SHOWNORMAL);
            }
        }
    }
}

// Show context menu
void ShowContextMenu() {
    POINT pt;
    GetCursorPos(&pt);
    SetForegroundWindow(g_hwnd);

    // Update autostart checkmark
    UINT autoStartState = IsAutoStartEnabled() ? MF_CHECKED : MF_UNCHECKED;
    CheckMenuItem(g_hMenu, ID_TRAY_AUTOSTART, MF_BYCOMMAND | autoStartState);

    TrackPopupMenu(g_hMenu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, pt.x, pt.y, 0, g_hwnd, NULL);
    PostMessage(g_hwnd, WM_NULL, 0, 0);
}

// Show about dialog
void ShowAboutDialog() {
    MessageBox(g_hwnd,
        L"Outlook to Tray\n\n"
        L"Minimizes Outlook to the system tray when closed.\n\n"
        L"Left-click tray icon to restore.\n"
        L"Right-click for menu.",
        L"About Outlook to Tray",
        MB_OK | MB_ICONINFORMATION);
}

// Create tray icon
bool CreateTrayIcon() {
    wchar_t dbg[256];

    // Verify window handle
    swprintf_s(dbg, L"CreateTrayIcon: hwnd=%p, IsWindow=%d", g_hwnd, IsWindow(g_hwnd));
    DebugMsg(dbg);

    if (!g_hwnd || !IsWindow(g_hwnd)) {
        MessageBox(NULL, L"Invalid window handle", L"Error", MB_ICONERROR);
        return false;
    }

    // Load gear icon from shell32.dll (icon index 316 is a gear on Windows 10/11)
    g_hIcon = ExtractIconW(g_hInstance, L"shell32.dll", 316);
    if (!g_hIcon || g_hIcon == (HICON)1) {
        // Fallback: try imageres.dll (index 109 = gear)
        g_hIcon = ExtractIconW(g_hInstance, L"imageres.dll", 109);
    }
    if (!g_hIcon || g_hIcon == (HICON)1) {
        // Final fallback: stock icon
        g_hIcon = LoadIcon(NULL, IDI_APPLICATION);
    }

    swprintf_s(dbg, L"LoadIcon result: %p", g_hIcon);
    DebugMsg(dbg);

    if (!g_hIcon) {
        MessageBox(NULL, L"Failed to load icon", L"Error", MB_ICONERROR);
        return false;
    }

    // Initialize NOTIFYICONDATA
    ZeroMemory(&g_nid, sizeof(g_nid));
    g_nid.cbSize = sizeof(NOTIFYICONDATA);
    g_nid.hWnd = g_hwnd;
    g_nid.uID = ID_TRAY_ICON;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = g_hIcon;
    lstrcpyW(g_nid.szTip, L"Outlook to Tray");

    swprintf_s(dbg, L"NOTIFYICONDATA: cbSize=%u, hWnd=%p, uID=%u",
               g_nid.cbSize, g_nid.hWnd, g_nid.uID);
    DebugMsg(dbg);

    // Try to add the icon
    BOOL result = Shell_NotifyIconW(NIM_ADD, &g_nid);
    swprintf_s(dbg, L"Shell_NotifyIconW result: %d", result);
    DebugMsg(dbg);

    if (!result) {
        // Retry a few times
        for (int i = 0; i < 3 && !result; i++) {
            Sleep(1000);
            result = Shell_NotifyIconW(NIM_ADD, &g_nid);
            swprintf_s(dbg, L"Retry %d: result=%d", i + 1, result);
            DebugMsg(dbg);
        }
    }

    if (!result) {
        MessageBox(NULL, L"Shell_NotifyIcon failed after retries.\n"
                        L"Try restarting Windows Explorer.",
                   L"Outlook to Tray", MB_ICONERROR);
        return false;
    }

    DebugMsg(L"Tray icon created successfully!");
    return true;
}

// Create context menu
void CreateContextMenu() {
    g_hMenu = CreatePopupMenu();
    AppendMenu(g_hMenu, MF_STRING, ID_TRAY_RESTORE, L"Restore Outlook");
    AppendMenu(g_hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(g_hMenu, MF_STRING, ID_TRAY_AUTOSTART, L"Run at Startup");
    AppendMenu(g_hMenu, MF_STRING, ID_TRAY_ABOUT, L"About");
    AppendMenu(g_hMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(g_hMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");
}

// Load DLL and get function pointers
bool LoadHookDll() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileName(NULL, exePath, MAX_PATH);

    std::wstring dllPath = exePath;
    size_t pos = dllPath.find_last_of(L'\\');
    if (pos != std::wstring::npos) {
        dllPath = dllPath.substr(0, pos + 1) + L"OutlookToTray.dll";
    }

    DebugMsg(L"Loading DLL...");
    DebugMsg(dllPath.c_str());

    g_hDll = LoadLibrary(dllPath.c_str());
    if (!g_hDll) {
        DWORD err = GetLastError();
        wchar_t buf[256];
        swprintf_s(buf, L"Failed to load DLL. Error: %lu", err);
        MessageBox(NULL, buf, L"Outlook to Tray", MB_ICONERROR);
        return false;
    }

    g_InstallHook = (InstallHookProc)GetProcAddress(g_hDll, "InstallHook");
    g_UninstallHook = (UninstallHookProc)GetProcAddress(g_hDll, "UninstallHook");
    g_GetHiddenOutlookWindow = (GetHiddenOutlookWindowProc)GetProcAddress(g_hDll, "GetHiddenOutlookWindow");
    g_ClearHiddenWindow = (ClearHiddenWindowProc)GetProcAddress(g_hDll, "ClearHiddenWindow");

    if (!g_InstallHook || !g_UninstallHook || !g_GetHiddenOutlookWindow) {
        MessageBox(NULL, L"DLL missing required functions", L"Outlook to Tray", MB_ICONERROR);
        FreeLibrary(g_hDll);
        g_hDll = NULL;
        return false;
    }

    DebugMsg(L"DLL loaded successfully");
    return true;
}

// Background thread: monitor for Outlook process
void MonitorOutlook() {
    DebugMsg(L"Monitor thread started");
    DWORD lastPid = 0;

    // Install hook immediately (pass DLL's module handle)
    if (g_InstallHook && g_hDll) {
        DebugMsg(L"Installing hook...");
        BOOL result = g_InstallHook(g_hDll);
        if (result) {
            DebugMsg(L"Hook installed successfully");
        } else {
            DebugMsg(L"Hook installation FAILED");
        }
    }

    while (g_running) {
        DWORD pid = FindProcessId(L"olk.exe");

        if (pid != 0 && lastPid == 0) {
            DebugMsg(L"Outlook detected");
        }
        else if (pid == 0 && lastPid != 0) {
            DebugMsg(L"Outlook closed");
        }

        lastPid = pid;
        Sleep(500);
    }

    // Cleanup
    if (g_UninstallHook) {
        DebugMsg(L"Removing hook...");
        g_UninstallHook();
    }
    DebugMsg(L"Monitor thread exiting");
}

#define WM_INITTRAY (WM_USER + 200)

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        DebugMsg(L"WM_CREATE");
        CreateContextMenu();
        // Defer tray icon creation
        PostMessage(hwnd, WM_INITTRAY, 0, 0);
        return 0;

    case WM_INITTRAY:
        DebugMsg(L"WM_INITTRAY - creating tray icon");
        if (!CreateTrayIcon()) {
            MessageBox(NULL, L"Could not create tray icon. Exiting.", L"Error", MB_ICONERROR);
            DestroyWindow(hwnd);
        }
        return 0;

    case WM_TRAYICON:
        // lParam contains the mouse message
        if (lParam == WM_LBUTTONUP || lParam == WM_LBUTTONDBLCLK) {
            RestoreOutlookWindow();
        }
        else if (lParam == WM_RBUTTONUP) {
            ShowContextMenu();
        }
        return 0;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_TRAY_RESTORE:
            RestoreOutlookWindow();
            break;
        case ID_TRAY_AUTOSTART:
            ToggleAutoStart();
            break;
        case ID_TRAY_ABOUT:
            ShowAboutDialog();
            break;
        case ID_TRAY_EXIT:
            DestroyWindow(hwnd);
            break;
        }
        return 0;

    case WM_DESTROY:
        DebugMsg(L"WM_DESTROY");
        g_running = false;
        Shell_NotifyIcon(NIM_DELETE, &g_nid);
        if (g_hMenu) DestroyMenu(g_hMenu);
        PostQuitMessage(0);
        return 0;

    default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}

// Entry point
int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                     LPSTR lpCmdLine, int nCmdShow) {
    DebugMsg(L"=== Outlook to Tray Starting ===");

    // Single instance check
    HANDLE hMutex = CreateMutex(NULL, TRUE, L"OutlookToTrayMutex");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBox(NULL, L"Outlook to Tray is already running.",
            L"Outlook to Tray", MB_OK | MB_ICONINFORMATION);
        CloseHandle(hMutex);
        return 0;
    }

    g_hInstance = hInstance;

    // Load hook DLL first
    if (!LoadHookDll()) {
        CloseHandle(hMutex);
        return 1;
    }

    // Register window class
    WNDCLASSEX wcex = {};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.lpfnWndProc = WindowProc;
    wcex.hInstance = hInstance;
    wcex.lpszClassName = L"OutlookToTrayClass";

    if (!RegisterClassEx(&wcex)) {
        DebugMsg(L"RegisterClassEx failed");
        FreeLibrary(g_hDll);
        CloseHandle(hMutex);
        return 1;
    }

    // Create a regular hidden window (not message-only)
    g_hwnd = CreateWindowEx(
        0,
        L"OutlookToTrayClass",
        L"Outlook to Tray",
        WS_OVERLAPPED,  // Regular window style
        0, 0, 0, 0,
        NULL,           // No parent (not HWND_MESSAGE)
        NULL,
        hInstance,
        NULL
    );

    if (!g_hwnd) {
        DWORD err = GetLastError();
        wchar_t buf[100];
        swprintf_s(buf, L"CreateWindowEx failed: %lu", err);
        DebugMsg(buf);
        MessageBox(NULL, buf, L"Error", MB_ICONERROR);
        FreeLibrary(g_hDll);
        CloseHandle(hMutex);
        return 1;
    }

    DebugMsg(L"Window created, starting monitor thread");

    // Start monitor thread
    std::thread monitorThread(MonitorOutlook);
    monitorThread.detach();

    DebugMsg(L"Entering message loop");

    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    DebugMsg(L"Exiting");

    // Cleanup
    if (g_hDll) {
        FreeLibrary(g_hDll);
    }
    CloseHandle(hMutex);

    return (int)msg.wParam;
}
