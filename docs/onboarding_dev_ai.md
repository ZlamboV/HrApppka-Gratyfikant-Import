# Developer & AI Agent Onboarding Guide

**Solution**: HrAppka Import Pracowników (v1.0.8.20)  
**Target Framework**: `.NET 8.0` (`net8.0-windows10.0.17763.0`)  
**SDK Dependency**: InsERT Gratyfikant Nexo Pro SDK 62.0.0.9537  
**Host Application**: InsERT Gratyfikant Nexo Pro (`Gratyfikant nexo.exe`)  
**Server Instance**: `100.127.240.115\INSERTGT`

---

## Table of Contents
1. [Introduction & Purpose](#1-introduction--purpose)
2. [Prerequisites & System Requirements](#2-prerequisites--system-requirements)
3. [Environment Configuration & Assembly Discovery](#3-environment-configuration--assembly-discovery)
4. [Step-by-Step Compilation & Build Guide](#4-step-by-step-compilation--build-guide)
5. [Testing & Verification Framework](#5-testing--verification-framework)
   - [5.1 Test Suite Organization (Tiers 1–4)](#51-test-suite-organization-tiers-14)
   - [5.2 Test Execution Commands](#52-test-execution-commands)
6. [Hot-Reload & Deployment Workflows](#6-hot-reload--deployment-workflows)
7. [Systematic 7-Scenario Troubleshooting Matrix](#7-systematic-7-scenario-troubleshooting-matrix)
   - [Scenario 1: Missing or Unresolvable SDK DLLs](#scenario-1-missing-or-unresolvable-sdk-dlls)
   - [Scenario 2: SQL Server Connection & Authentication Failures](#scenario-2-sql-server-connection--authentication-failures)
   - [Scenario 3: COM / Sfera Registration Failures](#scenario-3-com--sfera-registration-failures)
   - [Scenario 4: File Lock on Plugin DLL During Synchronization](#scenario-4-file-lock-on-plugin-dll-during-synchronization)
   - [Scenario 5: MiniExcel / CSV Parsing Format Mismatches](#scenario-5-miniexcel--csv-parsing-format-mismatches)
   - [Scenario 6: ZUS KEDU Schema Limit Violations](#scenario-6-zus-kedu-schema-limit-violations)
   - [Scenario 7: WPF Modal Window Hiding Behind Gratyfikant Main Window](#scenario-7-wpf-modal-window-hiding-behind-gratyfikant-main-window)
8. [AI Agent Coding Conventions & Safety Rules](#8-ai-agent-coding-conventions--safety-rules)

---

## 1. Introduction & Purpose

This guide establishes standard operating procedures for both **human software engineers** and **autonomous AI coding agents** developing, testing, maintaining, and deploying the **HrAppka Import Pracowników** plugin.

The codebase interfaces closely with the proprietary COM/Sfera object model and Entity Framework context of InsERT Gratyfikant Nexo Pro. Because of strict Mox DataDomain boundaries, STA UI thread constraints, and external SQL Server registration manifests, development requires careful adherence to the environment configurations and workflows detailed below.

---

## 2. Prerequisites & System Requirements

Before attempting to restore, compile, or run tests, ensure your host machine satisfies the following prerequisites:

| Component | Required Version | Verification Command | Description / Purpose |
|---|---|---|---|
| **Operating System** | Windows 10 (1809+) or Windows 11 (x64) | `[System.Environment]::OSVersion` | Required for WPF Desktop, Win32 window handles, and InsERT Sfera runtime COM bindings. |
| **.NET SDK** | .NET 8.0 SDK (8.0.100 or higher) | `dotnet --version` | Main development runtime. Must include `net8.0-windows10.0.17763.0` desktop targeting pack. |
| **Windows 10 SDK** | Build 10.0.17763.0 or higher | Installed via Visual Studio Build Tools | Required by WPF project configuration `<TargetFramework>net8.0-windows10.0.17763.0</TargetFramework>`. |
| **InsERT Nexo SDK** | Version 62.0.0.9537 (or 61.0.4.9400) | Check path: `C:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\` | Contains official SDK headers, XML documentation, and `PackageAndInstallerBuilder.exe`. |
| **Gratyfikant Nexo Client** | Version 62.0.0+ | Check path: `%LOCALAPPDATA%\InsERT\Deployments\Nexo\` | Local deployment folder where Gratyfikant Nexo Pro unpacks assemblies and executes the plugin. |
| **Database Server** | MS SQL Server 2017+ | `Test-NetConnection 100.127.240.115 -Port 1433` | Dedicated SQL Server hosting `InsERT_Launcher` and tenant databases (`Nexo_Test_1`, etc.). |
| **Python** | Python 3.8+ | `python --version` | Required to execute build helper scripts such as `scripts/bump_version.py`. |

---

## 3. Environment Configuration & Assembly Discovery

The solution links against two types of libraries:
1. **NuGet Packages**: Restored from nuget.org (e.g. `MiniExcel`, `xunit`, `Microsoft.NET.Test.Sdk`).
2. **Proprietary InsERT Mox/Moria DLLs**: Proprietary binaries not distributed via public NuGet.

### Assembly Reference Paths in `.csproj`
In `HrAppka Import Pracowników.csproj`, reference assemblies are dynamically sourced via the MSBuild property `<nexoSdkBinPath>`:

```xml
<PropertyGroup>
  <TargetFramework>net8.0-windows10.0.17763.0</TargetFramework>
  <OutputType>Exe</OutputType>
  <UseWPF>true</UseWPF>
  <nexoSdkBinPath>C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo\KLINEX SP_Z_O_Oa4bfb87390e94eea9\Binaries\</nexoSdkBinPath>
  <nexoSdkNarzedziaPath>C:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\Narzedzia\</nexoSdkNarzedziaPath>
</PropertyGroup>
```

### Dynamic Assembly Resolution at Runtime
When executing outside of the full Nexo client (e.g. during test runner execution or CLI diagnostic runs in `Program.cs`), an assembly resolving event handler resolves assemblies from multiple fallback paths:

```csharp
// Program.cs - Dynamic Assembly Load Hook
AssemblyLoadContext.Default.Resolving += (context, name) =>
{
    string assemblyName = name.Name;
    string[] searchDirs = new[]
    {
        @"C:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\Bin\",
        @"C:\ NEXO SDK SFERA\nexoSDK_61.0.4.9400\Bin\",
        @"C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo\KLINEX SP_Z_O_Oa4bfb87390e94eea9\Binaries\",
        AppDomain.CurrentDomain.BaseDirectory
    };

    foreach (var dir in searchDirs)
    {
        if (Directory.Exists(dir))
        {
            string candidate = Path.Combine(dir, assemblyName + ".dll");
            if (File.Exists(candidate))
            {
                try { return context.LoadFromAssemblyPath(candidate); } catch { }
            }
        }
    }
    return null;
};
```

---

## 4. Step-by-Step Compilation & Build Guide

### Step 1: Restore Dependencies
From the repository root or project directory, restore NuGet dependencies:
```powershell
cd "C:\ NEXO SDK SFERA\HrAppka Import Pracowników"
dotnet restore "HrAppka Import Pracowników.csproj"
```

### Step 2: Build Debug Configuration
Debug configuration is used for active development, local unit tests, and hot-syncing binaries to the client:
```powershell
dotnet build "HrAppka Import Pracowników.csproj" -c Debug --no-incremental
```

### Step 3: Build Release Configuration
Release configuration optimizes binaries and prepares assemblies for final `.mpkg` packaging:
```powershell
dotnet build "HrAppka Import Pracowników.csproj" -c Release --no-incremental
```

### Post-Build Action:
The `.csproj` contains a post-build target that automatically builds the installer package if `PackageAndInstallerBuilder.exe` is present:
```xml
<Target Name="PostBuild" AfterTargets="PostBuildEvent">
  <Exec Command="if exist &quot;$(nexoSdkNarzedziaPath)PackageAndInstallerBuilder.exe&quot; (call &quot;$(nexoSdkNarzedziaPath)PackageAndInstallerBuilder.exe&quot; /PlikKonfiguracji:Konfiguracja.xml /KatalogDocelowy:&quot;$(TargetDir)\&quot;) else (echo PackageAndInstallerBuilder not found)" />
</Target>
```

---

## 5. Testing & Verification Framework

The solution includes an enterprise test suite located in `C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj` covering unit, boundary, combination, and simulated end-to-end scenarios.

### 5.1 Test Suite Organization (Tiers 1–4)

```
tests/
├── Common/                      # Test doubles, in-memory repositories, Sfera mock helpers
├── Tier1_FeatureCoverage/       # 40 core functional tests (parsing, normalizers, converters)
├── Tier2_BoundaryCorner/        # 40 boundary tests (ZUS limits, leap years, exotic characters)
├── Tier3_CrossFeature/          # 6 combination tests (contract + ZUS + address enrichment)
└── Tier4_RealWorldScenarios/    # 4 end-to-end import simulations with full row datasets
```

### 5.2 Test Execution Commands

Execute tests using `dotnet test` with filter expressions targeting specific tiers:

```powershell
# Run the complete test suite
dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" --logger "console;verbosity=normal"

# Run Tier 1: Feature Coverage (40 tests)
dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" --filter "FullyQualifiedName~Tier1_FeatureCoverage"

# Run Tier 2: Boundary & Corner Cases (40 tests)
dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" --filter "FullyQualifiedName~Tier2_BoundaryCorner"

# Run Tier 3: Cross-Feature Combinations (6 tests)
dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" --filter "FullyQualifiedName~Tier3_CrossFeature"

# Run Tier 4: Real-World Ingestion Scenarios (4 tests)
dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" --filter "FullyQualifiedName~Tier4_RealWorldScenarios"
```

---

## 6. Hot-Reload & Deployment Workflows

When developing locally, avoid repeating full package builds and SQL database manifest registrations for every small code modification. Use the rapid **QuickSync** cycle:

```
[Edit C# / XAML] ──► [dotnet build -c Debug] ──► [sync_dll_to_all_deployments.ps1 -KillRunningNexo] ──► [Test in Nexo]
```

### Rapid QuickSync Command:
```powershell
cd "C:\ NEXO SDK SFERA\HrAppka Import Pracowników"
dotnet build -c Debug; powershell -ExecutionPolicy Bypass -File "scripts\sync_dll_to_all_deployments.ps1" -KillRunningNexo
```

---

## 7. Systematic 7-Scenario Troubleshooting Matrix

The following matrix documents the seven most common failure modes encountered by engineers and AI agents, their underlying root causes, and concrete remediation steps.

```
+---------------------------------------------------------------------------------------------------+
|                                Troubleshooting Scenarios Matrix                                    |
+----+---------------------------------------+------------------------------------------------------+
| #  | Scenario Category                     | Primary Diagnostic Symptom                           |
+----+---------------------------------------+------------------------------------------------------+
| 1  | Missing or Unresolvable SDK DLLs      | FileNotFoundException: InsERT.Moria.Sfera.dll        |
| 2  | SQL Server Connection / Auth Failure  | SqlException: Login failed for user 'sa'             |
| 3  | COM / Sfera Registration Failure      | Package does not appear in Gratyfikant menu          |
| 4  | File Lock on Plugin DLL During Sync   | IOException: Process cannot access file              |
| 5  | MiniExcel / CSV Format Mismatches     | Header row mismatch or multiline parsing error       |
| 6  | ZUS KEDU Schema Limit Violations      | Validation failed: string length exceeds 30 chars    |
| 7  | WPF Window Hiding Behind Gratyfikant  | UI freezes or modal opens behind main Nexo window    |
+----+---------------------------------------+------------------------------------------------------+
```

---

### Scenario 1: Missing or Unresolvable SDK DLLs

- **Symptom**:
  ```text
  System.IO.FileNotFoundException: Could not load file or assembly 'InsERT.Moria.Sfera, Version=62.0.0.0...' 
  or System.IO.FileNotFoundException: Could not load file or assembly 'InsERT.Mox.Core.dll'.
  ```
- **Root Cause**:
  The `<nexoSdkBinPath>` property in `HrAppka Import Pracowników.csproj` points to a path that does not exist on the current machine, or the local Nexo client has been updated to a different directory.
- **Remediation**:
  1. Inspect where Nexo binaries reside on the local machine:
     ```powershell
     Get-ChildItem -Path "$env:LOCALAPPDATA\InsERT\Deployments\Nexo" -Filter "InsERT.Moria.Sfera.dll" -Recurse
     ```
  2. Open `HrAppka Import Pracowników.csproj` and update `<nexoSdkBinPath>` to the valid directory:
     ```xml
     <nexoSdkBinPath>C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo\<YourDeployment>\Binaries\</nexoSdkBinPath>
     ```
  3. Ensure that `C:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\Bin\` contains official SDK references as fallback.

---

### Scenario 2: SQL Server Connection & Authentication Failures

- **Symptom**:
  ```text
  Microsoft.Data.SqlClient.SqlException (0x80131904): A network-related or instance-specific error occurred 
  while establishing a connection to SQL Server. Cannot open database "InsERT_Launcher". Login failed for user 'sa'.
  ```
- **Root Cause**:
  SQL Server instance `100.127.240.115\INSERTGT` is unreachable over the network, SQL Server Browser service is stopped, or credentials in `ParametryInstalacji.txt` are incorrect.
- **Remediation**:
  1. Test TCP port 1433 connectivity:
     ```powershell
     Test-NetConnection -ComputerName 100.127.240.115 -Port 1433
     ```
  2. If connection times out, verify that Tailscale/VPN or local network route to `100.127.240.115` is enabled.
  3. Verify connection configuration in `ParametryInstalacji.txt`:
     ```ini
     Serwer=100.127.240.115\INSERTGT
     UwierzytelnianieWindows=Nie
     Uzytkownik=sa
     Haslo=
     BazaDanych=Nexo_Test_1
     ```

---

### Scenario 3: COM / Sfera Registration Failures

- **Symptom**:
  The package is built and `install_to_target_dbs.ps1` runs without errors, but the operation `"Import Pracowników"` does not appear in the Gratyfikant Nexo Pro menu.
- **Root Cause**:
  1. `GrupaOperacji.cs` is missing the registration of the operation.
  2. The `DeploymentManifest` in `InsLauncher.InstalledProducts` on the target database has not been refreshed, or the local client cache has a corrupted manifest.
- **Remediation**:
  1. Check `GrupaOperacji.cs` to confirm `OperacjaImportuPracownikow20` is registered:
     ```csharp
     public IEnumerable<Operacja> Operacje { get; } = new Operacja[]
     {
         new OperacjaImportuPracownikow20()
     };
     ```
  2. Run `scripts/check_db_status.ps1` to verify if `<Package Identity="HrAppka_Import_Pracowników-1.0.8.20" ... />` is present in the XML manifest of the target database.
  3. Clear Nexo local client deployment cache if necessary by closing Nexo and re-running `scripts/install_to_target_dbs.ps1`.

---

### Scenario 4: File Lock on Plugin DLL During Synchronization

- **Symptom**:
  ```text
  System.IO.IOException: The process cannot access the file 
  'C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo\...\Binaries\HrAppka_Import_Pracowników.dll' 
  because it is being used by another process.
  ```
- **Root Cause**:
  `Gratyfikant nexo.exe` or `Nexo.exe` is currently running and has memory-mapped the plugin assembly.
- **Remediation**:
  Run `sync_dll_to_all_deployments.ps1` with the `-KillRunningNexo` switch:
  ```powershell
  powershell -ExecutionPolicy Bypass -File "scripts\sync_dll_to_all_deployments.ps1" -KillRunningNexo
  ```
  This immediately executes:
  ```powershell
  Get-Process | Where-Object { $_.ProcessName -like "*Gratyfikant*" -or $_.ProcessName -like "*Nexo*" } | Stop-Process -Force
  ```

---

### Scenario 5: MiniExcel / CSV Parsing Format Mismatches

- **Symptom**:
  Import fails during file load with:
  `InvalidDataException: Header column 'Nazwisko' not found` or multiline address strings in CSV fail to parse correctly.
- **Root Cause**:
  1. File is a CSV encoded with UTF-8 BOM, but standard comma delimiter was expected instead of semicolon.
  2. Address or notes contain multiline quotes that break standard line-by-line readers.
- **Remediation**:
  Always use `UniversalFileLoader.LoadFile(filePath)` instead of calling `MiniExcel.Query` directly. `UniversalFileLoader` checks extension:
  - If `.xlsx`: Uses `MiniExcel.Query<PracownikImportRow>(filePath)`.
  - If `.csv`: Delegates to `Rfc4180CsvReader.ReadRows(filePath)`, which handles RFC 4180 multiline cells, escapes quotes, detects semicolons, and strips UTF-8 BOM.

---

### Scenario 6: ZUS KEDU Schema Limit Violations

- **Symptom**:
  Saving a ZUS declaration throws validation exceptions:
  `Wartość pola 'Nazwisko' przekracza dopuszczalny limit 30 znaków` or `Brak numeru PESEL oraz dokumentu tożsamości`.
- **Root Cause**:
  ZUS KEDU XML schemas enforce strict character limits on personal data (e.g. max 30 chars for Surname, max 22 chars for First Name, max 30 chars for Street). Foreign employees frequently present combined double surnames or missing PESEL.
- **Remediation**:
  Ensure all data passes through `ZusSchemaLimits.Sanitize(...)` before being assigned to Moria entities:
  ```csharp
  // Services/ZusSchemaLimits.cs
  public static string Truncate(string? value, int maxLength)
  {
      if (string.IsNullOrWhiteSpace(value)) return string.Empty;
      string trimmed = value.Trim();
      return trimmed.Length <= maxLength ? trimmed : trimmed.Substring(0, maxLength);
  }
  ```
  If both PESEL and Passport/Identity document are missing, `PreAnalyzeRows` marks the row with `Status = "Brak PESEL i Paszportu"` and skips ZUS registration to prevent fatal transaction aborts.

---

### Scenario 7: WPF Modal Window Hiding Behind Gratyfikant Main Window

- **Symptom**:
  Opening the import preview freezes Gratyfikant Nexo Pro, and no modal window appears on the screen (the window is hidden behind the main application window or running on a non-STA background thread).
- **Root Cause**:
  Gratyfikant Nexo is a Win32/WPF hybrid application. Creating a WPF `Window` without explicitly linking its `Owner` handle causes Windows to position it behind the main window. Furthermore, instantiating WPF controls on a thread pool worker throws an STA thread exception.
- **Remediation**:
  Always display interactive dialogs via `ModalWindowHostingService.ShowModal(...)`:
  ```csharp
  // UI/Services/ModalWindowHostingService.cs
  bool? result = ModalWindowHostingService.ShowModal(() =>
  {
      var window = new ImportPreviewWindow();
      window.DataContext = viewModel;
      return window;
  });
  ```
  `ModalWindowHostingService`:
  1. Checks if `Application.Current.Dispatcher` is available and invokes on the STA thread.
  2. Sets `window.Owner = Application.Current.MainWindow`.
  3. If `MainWindow` is null, retrieves `Process.GetCurrentProcess().MainWindowHandle` and attaches it via `new WindowInteropHelper(window).Owner = hwnd`.

---

## 8. AI Agent Coding Conventions & Safety Rules

When autonomous AI coding agents implement fixes or extensions in this repository, they must adhere to the following rules:

1. **Source Code Preservation (Rule R4)**: Never delete or overwrite files without explicit authorization. Preserve existing comments and XML docstrings.
2. **Entity Framework Context Sponsorship**: When assigning entities (e.g. `Pracownik`, `OrganPodatkowy`, `SzablonUmowy`) to a Business Object (`IObiektBiznesowy`), always ensure the entity belongs to that BO's `ObjectContext` (`ModelDanychContainer`). Never cross-assign entities fetched from independent manager contexts.
3. **No Thread Blocking on STA**: Never run long I/O loops or database saves directly on the WPF Dispatcher. Use `Task.Run` in background workers and report progress to UI via `IProgress<T>`.
4. **Verified SDK APIs Only**: Do not invent or hallucinate InsERT Moria/Mox methods. Verify every method and property against decompiled DLL metadata in `C:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\Bin\`.
5. **Batch ZUS Declarations**: Never call `deklBO.Zapisz()` inside a per-row loop. Always use `ZusDeclarationBatchService` to assemble and commit declarations in a single consolidated transaction.

---
*Document author: Antigravity Technical Documentation Worker (`worker_doc_m2`)*  
*Last verified against codebase: November 2026*
