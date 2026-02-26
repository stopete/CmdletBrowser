# PowerShell Cmdlet Browser — C# / WPF

A Windows desktop application that replicates (and improves on) the original PowerShell WPF cmdlet browser script, converted to a proper C# .NET project.

---

<img width="1475" height="935" alt="image" src="https://github.com/user-attachments/assets/ee92ce78-5e83-4e5e-89e4-d0046c08559d" />


---

## Requirements

| Requirement | Version |
|---|---|
| OS | Windows 10 or 11 (64-bit) |
| .NET Framework | 4.8 (pre-installed on Windows 10+) |
| IDE | Visual Studio 2019/2022 **or** VS Code + C# Dev Kit |
| PowerShell | Windows PowerShell 5.1 (the `System.Management.Automation` reference) |

> **Note:** The app references `System.Management.Automation.dll` from the GAC (Windows PowerShell 5.1).  
> If the GAC path differs on your machine, update the `<HintPath>` in `CmdletBrowser.csproj`.

---

## Project Structure

```
CmdletBrowser.sln                  ← Visual Studio solution
CmdletBrowser/
  CmdletBrowser.csproj             ← Project file (.NET 4.8 WPF)
  App.xaml / App.xaml.cs           ← WPF Application entry point
  MainWindow.xaml                  ← Full dark-theme WPF UI
  MainWindow.xaml.cs               ← All application logic
  Properties/
    AssemblyInfo.cs                ← Assembly metadata
```

---

## Build & Run

### Visual Studio
1. Open `CmdletBrowser.sln`.
2. Set the build target to **x64**.
3. Press **F5** (Debug) or **Ctrl+F5** (Run without debugger).

### .NET CLI / MSBuild
```powershell
# From the repo root
dotnet build CmdletBrowser\CmdletBrowser.csproj -c Release
dotnet run   --project CmdletBrowser\CmdletBrowser.csproj
```

---

## Features

| Feature | Details |
|---|---|
| **Module tree** | Left-side tree groups all loaded modules with count |
| **Command list** | Filtered list with Name / Module / Type columns |
| **Live search** | Type to instantly filter commands by name |
| **Include Functions / Aliases** | Checkboxes to expand beyond Cmdlets |
| **Synopsis** | One-line description from `Get-Help` |
| **Syntax** | All parameter sets via `Get-Command -Syntax` |
| **Parameters grid** | Name, Type, Required, Position, Pipeline, Aliases |
| **Examples** | Code + remarks from `Get-Help` |
| **Copy Name / Syntax** | One-click clipboard copy |
| **Show Help (Window)** | Opens `Get-Help -ShowWindow` (WinPS 5.1) |
| **Open Online Help** | Opens Microsoft Docs search in browser |
| **Export CSV** | Saves the current filtered list to CSV |
| **Dark theme** | VS Code-inspired dark colour scheme |

---

## Troubleshooting

**`System.Management.Automation` not found**  
Update the `<HintPath>` in `CmdletBrowser.csproj` to match your system:
```
C:\Windows\assembly\GAC_MSIL\System.Management.Automation\1.0.0.0__31bf3856ad364e35\System.Management.Automation.dll
```
Or find it with:
```powershell
[System.Reflection.Assembly]::LoadWithPartialName('System.Management.Automation').Location
```

**Help shows "No local synopsis available"**  
Run `Update-Help -ErrorAction SilentlyContinue` in an elevated PowerShell window to download help files.

**`Get-Help -ShowWindow` unavailable**  
This feature requires Windows PowerShell 5.1 as the host process. Use **Open Online Help** as a fallback.
