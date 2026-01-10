# Mouse Wheel Scroll Reverse

[中文](./README_CN.md)

## 1. Overview

This tool reverses the mouse wheel scroll direction by modifying the Windows Registry via PowerShell. The `.ps1` script can be converted into a standalone `.exe` executable using **PS2EXE**.

![img1](./docs/screenshot1.png "Screenshot")

## 2. Instructions

### 2.1 Running the Tool

There are two ways to run this tool: via a batch file (`.bat`) or as a compiled executable (`.exe`).

- **Batch method**
  Simply run [run.bat](src/run.bat) to launch the script directly.

- **Executable method**
  Use the [build.bat](src/build.bat) batch script to convert the `.ps1` file into an `.exe` executable, then double-click to run it.

### 2.2 Mouse Database Maintenance

This project uses [vendors.json](src/vendors.json) to maintain mappings of mouse vendor IDs (`VID`) and product IDs (`PID`).

You can extend this database yourself using resources like:

- <http://www.linux-usb.org/usb.ids>
- <https://devicehunt.com/>

## 3. Appendix

### 3.1 Installing `PS2EXE`

Open **PowerShell as Administrator** and run the following command to install the module:

```powershell
# Install the PS2EXE module
Install-Module -Name PS2EXE -Force
# Verify installation
Get-Module -ListAvailable PS2EXE
```

> If you see an error like *"Installation not allowed because the repository is untrusted,"* first run:
> `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` and press `Y` to confirm.

### 3.2 Packaging a PowerShell Script into EXE

In PowerShell, navigate to your script’s directory and run:

```powershell
# Basic syntax: ps2exe <source_script.ps1> <output.exe>
ps2exe .\your_script.ps1 .\output.exe
```

### 3.3 Advanced Options (Optional)

You can customize the EXE’s icon, window style, required privileges, etc.:

```powershell
ps2exe .\your_script.ps1 .\output.exe `
-Icon .\custom.ico `          # Custom application icon
-WindowStyle Hidden `         # Hide console window (run in background)
-AdminRequired `              # Require administrator privileges
-NoConsole `                  # No console window at all (ideal for GUI apps)
```

### 3.4 Notes

- The generated EXE **requires .NET Framework**, which is typically pre-installed on Windows.
- If the script accesses the file system or registry, the EXE will run with the same permissions as the user who launched it.
