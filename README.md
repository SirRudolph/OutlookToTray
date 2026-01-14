# Outlook to Tray

Minimizes the new Outlook (olk.exe) to the system tray when you close the window, instead of exiting the application.

## Features

- Intercepts window close and hides Outlook instead of closing it
- System tray icon with right-click menu
- Left-click tray icon to restore Outlook
- Option to run at Windows startup
- Lightweight and runs in the background

## Requirements

- Windows 10/11 (64-bit)
- New Outlook app (olk.exe)
- MSYS2 with MinGW-w64 (for building)

## Building

### Prerequisites

1. Install [MSYS2](https://www.msys2.org/)
2. Open MSYS2 terminal and install the toolchain:
   ```bash
   pacman -S mingw-w64-x86_64-gcc mingw-w64-x86_64-make
   ```
3. Add `C:\msys64\mingw64\bin` to your Windows PATH

### Build Steps

1. Clone the repository:
   ```bash
   git clone https://github.com/YOUR_USERNAME/OutlookToTray.git
   cd OutlookToTray
   ```

2. Run the build script:
   ```bash
   build.bat
   ```

3. The output files will be in the `bin/` folder:
   - `OutlookToTray.exe` - Main application
   - `OutlookToTray.dll` - Hook library

## Usage

1. Copy both `OutlookToTray.exe` and `OutlookToTray.dll` to the same folder
2. Run `OutlookToTray.exe`
3. A gear icon will appear in your system tray
4. Open Outlook and use it normally
5. When you close Outlook's window, it will hide to the tray instead of exiting
6. Left-click the tray icon to restore Outlook
7. Right-click the tray icon for options:
   - **Restore Outlook** - Show the hidden window
   - **Run at Startup** - Toggle automatic startup with Windows
   - **About** - Version information
   - **Exit** - Close the application

## How It Works

The application uses a Windows hook (WH_CALLWNDPROC) to intercept window messages. When Outlook's main window receives a WM_CLOSE message, the hook hides the window instead of allowing it to close. A memory-mapped file is used for cross-process communication between the hook DLL and the main application.

## Project Structure

```
OutlookToTray/
├── OutlookToTray.Dll/           # Hook DLL
│   └── OutlookToTray.Dll.cpp    # Hook implementation
├── OutlookToTray.Exe/           # Main application
│   ├── OutlookToTray.Exe.cpp    # Tray app implementation
│   ├── resource.h               # Resource definitions
│   └── OutlookToTray.rc         # Resource script
├── build.bat                    # Build script for MinGW
├── Makefile                     # Alternative Makefile
└── OutlookToTray.sln            # Visual Studio solution (optional)
```

## License

MIT License - feel free to use and modify as you like.
