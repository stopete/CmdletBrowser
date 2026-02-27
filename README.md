
# CmdletBrowser

A modern Windows desktop app to **explore, search, and inspect PowerShell commands** through a fast, user‑friendly GUI. CmdletBrowser hosts a PowerShell runspace internally so you can browse commands, view syntax, parameters, and examples—without opening a console.

---

##
 📸 Screenshot
<img width="1475" height="935" alt="Screenshot 2026-02-26 153801" src="https://github.com/user-attachments/assets/c43eb984-199d-4195-b4ef-fcb2acfe53bc" />


> 

---

## ✨ Features

- **Browse all commands** – cmdlets, functions, and aliases, grouped by module (with a special group for commands that have no module).
- **Instant filtering** – search by command name, type, and module.
- **Deep help view** – shows *Synopsis*, *Syntax* (all parameter sets), and *Examples* (when available) by calling `Get-Help` and `Get-Command` under the hood.  
  > Tip: `Get-Help -Full` returns the most complete help object, including examples and parameter details. citeturn11search61
- **Parameter grid** – name, type, required/optional, position, pipeline input support, and aliases.
- **One‑click actions** – copy name, copy syntax, open help window, open online docs.
- **Export to CSV** – export the currently filtered command list.
- **Responsive UI** – background runspaces for long operations keep the app snappy.

---

## 🧰 Tech Stack

- **.NET**: `net8.0-windows` WinForms
- **PowerShell Hosting**: [`Microsoft.PowerShell.SDK`](https://www.nuget.org/packages/Microsoft.PowerShell.SDK/) (PowerShell 7 runtime and APIs) – the SDK targets modern .NET TFMs. 

---

## 🚀 Getting Started

### Prerequisites
- Windows 10/11
- Visual Studio 2022 17.x with .NET 8 SDK

### Clone & Open
```bash
# using Git
git clone https://github.com/<your-org>/<your-repo>.git
cd <your-repo>
```
Open the solution in **Visual Studio 2022** and build.

### Restore NuGet packages (with Package Source Mapping enabled)
If you use **Package Source Mapping**, ensure `nuget.org` is mapped so core packages like `System.Collections.Immutable` can restore. Example `nuget.config` at the solution root:

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
  </packageSources>
  <packageSourceMapping>
    <packageSource key="nuget.org">
      <package pattern="*" />
    </packageSource>
  </packageSourceMapping>
</configuration>
```
NuGet only searches sources explicitly mapped for a package when Package Source Mapping is enabled; unmapped sources are **not considered**. 

### First‑run: make Examples appear
`Get-Help` shows full examples only when local help files are installed. Run this once in **Windows PowerShell 5.1 (Run as Administrator)** to install updatable help for built‑in modules:

```powershell
Update-Help -Module * -UICulture en-US -Force -ErrorAction SilentlyContinue
```
- `Get-Help` uses local help files; otherwise it returns only basic, auto‑generated help. 
- `Update-Help` downloads and installs the newest help files; elevation is required on PowerShell 5.1.

> Offline/isolated machines: use `Save-Help` on a connected machine, then `Update-Help -SourcePath` on the target.

---

## 🖱️ How to Use

1. **Refresh** to load all available commands (cmdlets, functions, aliases).
2. **Filter** by typing in *Search* and/or selecting a module in the left tree.
3. Click a command to view **Synopsis**, **Syntax**, **Examples**, and the **Parameters** grid.
4. Use **Copy Name**, **Copy Syntax**, **Show Help (Window)**, and **Open Online Help** as needed.
5. Click **Export CSV** to save the current command list.

---

## 📁 Project Structure (high‑level)

```
src/
  CmdletBrowser/            # WinForms app
    MainForm.cs             # UI + event handlers
    PowerShell helpers      # runspace + Get-Command / Get-Help wrappers
images/
  ModuleBrowser.png         # screenshot used by README
```

---

## 🔧 Build & Packaging

- Build in **Release**: `Ctrl+Shift+B` (VS) or `dotnet build -c Release`.
- Optional single‑file publish:
  ```bash
  dotnet publish -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true
  ```

---

## ❓ Troubleshooting

- **Examples show “No examples available.”**  
  Install/update local help (`Update-Help`), and the app will display examples that modules actually provide. `Get-Help -Full` exposes the examples collection used by the app. 

- **Restore fails with “source(s) were not considered: nuget.org.”**  
  You have Package Source Mapping enabled but didn’t map the package IDs to nuget.org in `nuget.config`. Add a mapping (e.g., `*` → nuget.org). 

---

## 📜 License
Choose a license (e.g., MIT) and add `LICENSE` to the repo.

---

## 🙌 Credits
- PowerShell help behavior: [`Get-Help`](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/get-help) and [`Update-Help`](https://learn.microsoft.com/powershell/module/microsoft.powershell.core/update-help). 
- NuGet configuration and Package Source Mapping: `nuget.config` reference and package‑source mapping docs.

