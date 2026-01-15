/*
 * Outlook to Tray - Hook DLL
 * Intercepts Outlook window close and hides to tray instead
 * Uses memory-mapped file for cross-process communication
 */

#include <windows.h>
#include <psapi.h>
#include <tchar.h>
#include <commctrl.h>

#pragma comment(lib, "Psapi.lib")
#pragma comment(lib, "Comctl32.lib")

// Shared memory structure
struct SharedData {
    HWND hiddenWindow;
    BOOL subclassed;
    RECT originalRect;      // Original window position before hiding
    LONG originalExStyle;   // Original extended style (to restore taskbar visibility)
};

// Global state (per-process)
HINSTANCE g_hInstance = NULL;
HHOOK g_hook = NULL;
HANDLE g_hMapFile = NULL;
SharedData* g_pShared = NULL;

const wchar_t* SHARED_MEM_NAME = L"OutlookToTraySharedMem";

// Get or create shared memory
SharedData* GetSharedData() {
    if (g_pShared) return g_pShared;

    // Try to open existing
    g_hMapFile = OpenFileMappingW(FILE_MAP_ALL_ACCESS, FALSE, SHARED_MEM_NAME);

    if (!g_hMapFile) {
        // Create new
        g_hMapFile = CreateFileMappingW(
            INVALID_HANDLE_VALUE, NULL, PAGE_READWRITE, 0, sizeof(SharedData), SHARED_MEM_NAME);
    }

    if (g_hMapFile) {
        g_pShared = (SharedData*)MapViewOfFile(g_hMapFile, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(SharedData));
    }

    return g_pShared;
}

// Check if window belongs to olk.exe (new Outlook)
BOOL IsOutlookProcess(HWND hwnd) {
    DWORD processId;
    GetWindowThreadProcessId(hwnd, &processId);

    HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processId);
    if (hProcess) {
        TCHAR processName[MAX_PATH];
        if (GetModuleBaseName(hProcess, NULL, processName, sizeof(processName) / sizeof(TCHAR))) {
            CloseHandle(hProcess);
            return _tcsicmp(processName, L"olk.exe") == 0;
        }
        CloseHandle(hProcess);
    }
    return FALSE;
}

// Check if this is a main/top-level window
BOOL IsMainWindow(HWND hwnd) {
    return GetWindow(hwnd, GW_OWNER) == NULL && IsWindowVisible(hwnd);
}

// Subclass procedure to block WM_CLOSE
LRESULT CALLBACK SubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam,
                               UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    if (uMsg == WM_CLOSE) {
        SharedData* pData = GetSharedData();
        if (pData) {
            // Save original position
            GetWindowRect(hwnd, &pData->originalRect);

            // Save original extended style and hide from taskbar
            pData->originalExStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            SetWindowLong(hwnd, GWL_EXSTYLE,
                         (pData->originalExStyle | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW);

            // Move window off-screen instead of hiding it completely
            // This allows notifications to still work (SW_HIDE suppresses them)
            SetWindowPos(hwnd, NULL, -32000, -32000, 0, 0,
                         SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);

            pData->hiddenWindow = hwnd;
        }
        return 0;  // Block the close
    }
    else if (uMsg == WM_DESTROY) {
        SharedData* pData = GetSharedData();
        if (pData && pData->hiddenWindow == hwnd) {
            pData->hiddenWindow = NULL;
            pData->subclassed = FALSE;
        }
        RemoveWindowSubclass(hwnd, SubclassProc, 1);
    }
    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

// Main hook callback - runs in the target process
LRESULT CALLBACK CallWndProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0) {
        CWPSTRUCT* pCwp = (CWPSTRUCT*)lParam;

        // Check for Outlook windows
        if (IsOutlookProcess(pCwp->hwnd) && IsMainWindow(pCwp->hwnd)) {
            SharedData* pData = GetSharedData();

            // Subclass on first WM_CLOSE or when window becomes visible
            if (pCwp->message == WM_CLOSE ||
                (pCwp->message == WM_SHOWWINDOW && pCwp->wParam == TRUE)) {

                if (pData && !pData->subclassed) {
                    if (SetWindowSubclass(pCwp->hwnd, SubclassProc, 1, 0)) {
                        pData->subclassed = TRUE;
                    }
                }
            }
        }
    }
    return CallNextHookEx(g_hook, nCode, wParam, lParam);
}

// Exported: Install the hook
extern "C" __declspec(dllexport) BOOL InstallHook(HINSTANCE hInst) {
    if (g_hook != NULL) {
        return TRUE;  // Already installed
    }

    // Initialize shared memory
    SharedData* pData = GetSharedData();
    if (!pData) {
        return FALSE;
    }

    g_hook = SetWindowsHookExW(WH_CALLWNDPROC, CallWndProc, hInst, 0);
    return (g_hook != NULL);
}

// Exported: Remove the hook
extern "C" __declspec(dllexport) BOOL UninstallHook() {
    if (g_hook != NULL) {
        UnhookWindowsHookEx(g_hook);
        g_hook = NULL;

        SharedData* pData = GetSharedData();
        if (pData) {
            pData->subclassed = FALSE;
        }
        return TRUE;
    }
    return FALSE;
}

// Exported: Get handle of hidden Outlook window
extern "C" __declspec(dllexport) HWND GetHiddenOutlookWindow() {
    SharedData* pData = GetSharedData();
    if (pData) {
        return pData->hiddenWindow;
    }
    return NULL;
}

// Exported: Clear hidden window tracking
extern "C" __declspec(dllexport) void ClearHiddenWindow() {
    SharedData* pData = GetSharedData();
    if (pData) {
        pData->hiddenWindow = NULL;
    }
}

// Exported: Get original window rect (for restoring position)
extern "C" __declspec(dllexport) BOOL GetOriginalRect(RECT* pRect) {
    SharedData* pData = GetSharedData();
    if (pData && pRect) {
        *pRect = pData->originalRect;
        return TRUE;
    }
    return FALSE;
}

// Exported: Get original extended style (for restoring taskbar visibility)
extern "C" __declspec(dllexport) LONG GetOriginalExStyle() {
    SharedData* pData = GetSharedData();
    if (pData) {
        return pData->originalExStyle;
    }
    return 0;
}

// DLL entry point
BOOL APIENTRY DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved) {
    switch (fdwReason) {
    case DLL_PROCESS_ATTACH:
        g_hInstance = hinstDLL;
        DisableThreadLibraryCalls(hinstDLL);
        break;
    case DLL_PROCESS_DETACH:
        if (g_pShared) {
            UnmapViewOfFile(g_pShared);
            g_pShared = NULL;
        }
        if (g_hMapFile) {
            CloseHandle(g_hMapFile);
            g_hMapFile = NULL;
        }
        break;
    }
    return TRUE;
}
