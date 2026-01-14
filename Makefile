# Outlook to Tray - Makefile for MinGW-w64

CXX = g++
CXXFLAGS = -std=c++17 -Wall -O2 -DUNICODE -D_UNICODE
LDFLAGS_DLL = -shared -static-libgcc -static-libstdc++
LDFLAGS_EXE = -mwindows -static-libgcc -static-libstdc++

# Libraries
LIBS_DLL = -lpsapi -lcomctl32
LIBS_EXE = -lshell32 -lpsapi

# Output directory
OUTDIR = bin

# Targets
DLL = $(OUTDIR)/OutlookToTray.dll
EXE = $(OUTDIR)/OutlookToTray.exe

# Source files
DLL_SRC = OutlookToTray.Dll/OutlookToTray.Dll.cpp
EXE_SRC = OutlookToTray.Exe/OutlookToTray.Exe.cpp
EXE_RC = OutlookToTray.Exe/OutlookToTray.rc

.PHONY: all clean run

all: $(OUTDIR) $(DLL) $(EXE)
	@echo Build complete! Run with: ./bin/OutlookToTray.exe

$(OUTDIR):
	mkdir -p $(OUTDIR)

$(DLL): $(DLL_SRC)
	$(CXX) $(CXXFLAGS) $(LDFLAGS_DLL) -o $@ $< $(LIBS_DLL)
	@echo Built: $@

$(EXE): $(EXE_SRC)
	windres $(EXE_RC) -o $(OUTDIR)/resources.o
	$(CXX) $(CXXFLAGS) $(LDFLAGS_EXE) -o $@ $< $(OUTDIR)/resources.o $(LIBS_EXE)
	@rm -f $(OUTDIR)/resources.o
	@echo Built: $@

clean:
	rm -rf $(OUTDIR)

run: all
	cd $(OUTDIR) && ./OutlookToTray.exe
