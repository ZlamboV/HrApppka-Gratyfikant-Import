# Deployment & Automation Scripts Guide

**Solution**: HrAppka Import Pracowników (v1.0.8.20)  
**Target Framework**: `.NET 8.0` (`net8.0-windows10.0.17763.0`)  
**SDK Dependency**: InsERT Gratyfikant Nexo Pro SDK 62.0.0.9537  
**Host Application**: InsERT Gratyfikant Nexo Pro (`Gratyfikant nexo.exe`)  
**Server Instance**: `100.127.240.115\INSERTGT`

---

## Table of Contents
1. [Overview & Deployment Architecture](#1-overview--deployment-architecture)
2. [Target Database Topology & Infrastructure](#2-target-database-topology--infrastructure)
3. [Automation Scripts Deep Dive](#3-automation-scripts-deep-dive)
   - [3.1 build_and_package_release.ps1](#31-build_and_package_releaseps1)
   - [3.2 install_to_target_dbs.ps1](#32-install_to_target_dbsps1)
   - [3.3 sync_dll_to_all_deployments.ps1](#33-sync_dll_to_all_deploymentsps1)
   - [3.4 check_db_status.ps1](#34-check_db_statusps1)
   - [3.5 bump_version.py](#35-bump_versionpy)
4. [End-to-End Release & Deployment Workflows](#4-end-to-end-release--deployment-workflows)
   - [4.1 Full Production Deployment Pipeline](#41-full-production-deployment-pipeline)
   - [4.2 Rapid Developer Hot-Reload (QuickSync)](#42-rapid-developer-hot-reload-quicksync)
5. [Script Parameter & Operational Matrix](#5-script-parameter--operational-matrix)
6. [Troubleshooting & Diagnostic Commands](#6-troubleshooting--diagnostic-commands)

---

## 1. Overview & Deployment Architecture

The **HrAppka Import Pracowników** solution runs as an in-process native Sfera extension inside **InsERT Gratyfikant Nexo Pro**. Unlike standalone client applications or external services, Gratyfikant Nexo extensions are managed by the InsERT Mox Deployment Infrastructure:

```
+-------------------------------------------------------------------------------------------------------+
|                                    InsERT Deployment Architecture                                     |
+-------------------------------------------------------------------------------------------------------+

       [Developer Workstation / CI]
                    │
                    │ 1. Compile .NET 8 DLLs
                    ▼
     [bin\Debug or bin\Release Outputs]
      - HrAppka_Import_Pracowników.dll
      - HrAppka_Import_Pracowników.plugin
      - MiniExcel.dll
                    │
                    │ 2. PackageAndInstallerBuilder.exe (/PlikKonfiguracji:Konfiguracja.xml)
                    ▼
   [Pakiet/*.mpkg] & [Instalator/*.exe]
                    │
                    │ 3. Automated Launcher Registration (install_to_target_dbs.ps1)
                    ├───────────────────────────────────────────────────────┐
                    ▼                                                       ▼
      [SQL Server Instance]                                    [Local Client Workstation]
       100.127.240.115\INSERTGT                                 %LOCALAPPDATA%\InsERT\Deployments\Nexo\
        ├── Database: InsERT_Launcher                           └── [Deployment Folder]\Binaries\
        │    └── Binary Package Blob Storage                         ├── HrAppka_Import_Pracowników.dll
        └── Tenant Databases:                                        ├── HrAppka_Import_Pracowników.pdb
             ├── Nexo_Test_1                                         ├── HrAppka_Import_Pracowników.plugin
             ├── Nexo_KLINEX SP_Z_O_O                                └── MiniExcel.dll
             ├── Nexo_EUROSUPPORT_GROUP
             └── Nexo_EUROSUPPORT OUTSOURCING
                  └── InsLauncher.InstalledProducts
                       └── DeploymentManifest XML (Registration)
```

### Key Mechanisms:
1. **Packaging Manifest (`Konfiguracja.xml`)**: Defines the metadata, package identity (`HrAppka_Import_Pracowników-1.0.8.20`), payload files, and installer parameters.
2. **InsERT Mox Package (`.mpkg`)**: A compressed cabinet archive containing the compiled assemblies, manifest, and plugin definition file (`.plugin`).
3. **Launcher Database (`InsERT_Launcher`)**: The central database where package binaries are uploaded and managed on the SQL Server instance.
4. **Tenant Deployment Manifest (`InsLauncher.InstalledProducts`)**: Each tenant database contains a registration record with an XML deployment manifest (`DeploymentManifest`) specifying which packages are activated for that company.
5. **Client Local Deployments (`%LOCALAPPDATA%\InsERT\Deployments\Nexo\`)**: When Gratyfikant Nexo boots, it checks the database manifest against its local deployment folder. If matching binaries are present in the `Binaries\` folder, Nexo loads them directly into memory.

---

## 2. Target Database Topology & Infrastructure

The project interacts with a dedicated Microsoft SQL Server hosting multiple company databases (tenants) and the shared launcher system:

| Environment | Database Name | Description | Active Status |
|---|---|---|---|
| **SQL Host** | `100.127.240.115\INSERTGT` | Dedicated Microsoft SQL Server instance | Online (Port 1433) |
| **Launcher** | `InsERT_Launcher` | InsERT central catalog storing uploaded `.mpkg` packages | System Database |
| **Test Tenant** | `Nexo_Test_1` | Primary staging / test environment for Sfera import verification | Active (Safe for tests) |
| **Demo Tenant** | `Nexo_Demo_1` | Diagnostic test database used for manifest status queries | Active |
| **Legacy Test** | `Nexo_Test_GR` | Legacy staging database | Deprecated |
| **Production 1** | `Nexo_KLINEX SP_Z_O_O` | Klinex Sp. z o.o. operational tenant database | Active Production |
| **Production 2** | `Nexo_EUROSUPPORT_GROUP` | Eurosupport Group Sp. z o.o. operational tenant database | Active Production |
| **Production 3** | `Nexo_EUROSUPPORT OUTSOURCING` | Eurosupport Outsourcing Sp. z o.o. operational tenant database | Active Production |

### SQL Server Connection Parameters
- **Server Address**: `100.127.240.115\INSERTGT`
- **Authentication**: SQL Server Native Authentication (`UwierzytelnianieWindows=Nie`)
- **Default Database User**: `sa`
- **TCP Port**: `1433` (SQL Browser active for instance name resolution)
- **Encryption**: `TrustServerCertificate=True`

---

## 3. Automation Scripts Deep Dive

All production automation workflows are implemented as PowerShell and Python scripts located in `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\scripts\`.

```
scripts/
├── build_and_package_release.ps1   # Compiles and builds .mpkg and .exe packages
├── install_to_target_dbs.ps1       # Deploys packages to SQL Server & updates manifests
├── sync_dll_to_all_deployments.ps1 # Hot-reloads binaries into local client runtime
├── check_db_status.ps1             # Queries SQL database to verify registration
└── bump_version.py                 # Synchronizes version strings across .csproj & .xml
```

---

### 3.1 `build_and_package_release.ps1`

- **Location**: `scripts/build_and_package_release.ps1`
- **Purpose**: Automates the collection of compiled assemblies from `bin\Debug\`, prepares them in the project root, and invokes the InsERT SDK `PackageAndInstallerBuilder.exe` utility to produce the final `.mpkg` deployment archive and standalone `.exe` installer.

#### Workflow & Execution Steps:
1. **Safety Configuration**: Sets `$ErrorActionPreference = 'Stop'` to abort on any error.
2. **Project Path Resolution**: Discovers project root using wildcard match on `HrAppka Import Prac*`.
3. **Artifact Preparation**: Copies target files from `bin\Debug\` to the project root directory:
   - `*Import_Pracownik*.dll`
   - `*Import_Pracownik*.plugin`
   - `MiniExcel.dll`
4. **Environment Isolation**: Sets telemetry and AppData environment variables:
   ```powershell
   $env:APPDATA = "c:\ NEXO SDK SFERA"
   $env:GIT_CONFIG_GLOBAL = "c:\ NEXO SDK SFERA\.gitconfig"
   $env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
   $env:DOTNET_NOLOGO = "1"
   ```
5. **Builder Invocation**: Invokes the SDK packaging utility:
   ```powershell
   & "C:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\Narzedzia\PackageAndInstallerBuilder.exe" "/PlikKonfiguracji:Konfiguracja.xml"
   ```
6. **Artifact Verification**: Inspects `Pakiet\` for the newest `.mpkg` file and `Instalator\` for the newest `.exe` file. If the installer builder produced the executable in `binDir`, copies it to `Instalator\`.
7. **Console Confirmation**: Prints byte sizes and absolute paths for both artifacts.

#### Invocation:
```powershell
powershell -ExecutionPolicy Bypass -File "C:\ NEXO SDK SFERA\HrAppka Import Pracowników\scripts\build_and_package_release.ps1"
```

#### Expected Console Output:
```text
Copying build outputs to root for packaging from C:\ NEXO SDK SFERA\HrAppka Import Pracowników\bin\Debug...
Running PackageAndInstallerBuilder.exe...
>>> SUCCESS! Package created: C:\ NEXO SDK SFERA\HrAppka Import Pracowników\Pakiet\HrAppka_Import_Pracowników-1.0.8.20.mpkg (Size: 217160 bytes) <<<
>>> SUCCESS! Installer created: C:\ NEXO SDK SFERA\HrAppka Import Pracowników\Instalator\HrAppka_Import_Pracowników-1.0.8.20.exe (Size: 25176171 bytes) <<<
```

#### Exit Codes & Errors:
- Exits with `0` on success.
- Throws fatal terminating error if `PackageAndInstallerBuilder.exe` is missing or fails execution.

---

### 3.2 `install_to_target_dbs.ps1`

- **Location**: `scripts/install_to_target_dbs.ps1`
- **Purpose**: Deploys the generated `.mpkg` package to the SQL Server launcher database (`InsERT_Launcher`), attaches the package to the specified tenant databases, and triggers binary synchronization to local client folders.

#### Script Parameters:
| Parameter | Type | Default Value | Description |
|---|---|---|---|
| `-Databases` | `string[]` | `@("Nexo_Demo_1", "Nexo_KLINEX SP_Z_O_O", "Nexo_EUROSUPPORT OUTSOURCING", "Nexo_EUROSUPPORT_GROUP")` | Target SQL Server database names to register the package into. |
| `-Server` | `string` | `"100.127.240.115\INSERTGT"` | Target SQL Server network instance. |
| `-QuickSyncOnly` | `switch` | `false` | When enabled, skips SQL Server registration and only synchronizes binaries to client folders. |

#### Workflow & Execution Steps:
1. **QuickSync Check**: If `-QuickSyncOnly` is set, calls `sync_dll_to_all_deployments.ps1` and immediately returns.
2. **Package Discovery**: Locates the newest `.mpkg` file inside `Pakiet\`. Throws if no package exists.
3. **Iteration Over Target Databases**: For each database in `$Databases`:
   - Writes `ParametryInstalacji.txt` to both project root and `bin\Debug\`:
     ```ini
     Serwer=100.127.240.115\INSERTGT
     UwierzytelnianieWindows=Nie
     Uzytkownik=sa
     Haslo=
     BazaDanych=<DatabaseName>
     ```
   - Executes the package installer entry point:
     ```powershell
     dotnet "$dllCandidate" --install
     ```
   - `AutoInstaller.Install()` connects to `InsERT_Launcher`, checks if package identity exists, uploads the binary blob, connects to the tenant database, reads `DeploymentManifest`, appends the package identity, and saves the updated manifest.
   - Cleans up any background `dotnet` processes spawned during registration.
4. **Client Synchronization**: Calls `sync_dll_to_all_deployments.ps1 -KillRunningNexo` to push the new binaries directly into local Nexo deployments.

#### Invocation Examples:
```powershell
# Deploy to test database only
powershell -ExecutionPolicy Bypass -File "scripts\install_to_target_dbs.ps1" -Databases @("Nexo_Test_1")

# Deploy to all production databases
powershell -ExecutionPolicy Bypass -File "scripts\install_to_target_dbs.ps1" -Databases @("Nexo_KLINEX SP_Z_O_O", "Nexo_EUROSUPPORT_GROUP", "Nexo_EUROSUPPORT OUTSOURCING")

# Fast client-only hot reload without modifying SQL Server manifests
powershell -ExecutionPolicy Bypass -File "scripts\install_to_target_dbs.ps1" -QuickSyncOnly
```

#### Expected Console Output:
```text
Found MPKG release package: C:\ NEXO SDK SFERA\HrAppka Import Pracowników\Pakiet\HrAppka_Import_Pracowników-1.0.8.20.mpkg
==========================================================
=== Installing plugin package onto database 'Nexo_Test_1' on server '100.127.240.115\INSERTGT' ===
==========================================================
Executing installer for 'Nexo_Test_1' with DLL 'C:\ ... \HrAppka_Import_Pracowników.dll'...
Ładowanie pakietu...
Wysyłanie pakietu do bazy InsERT_Launcher...
Manifest został pomyślnie zaktualizowany.
>>> SUCCESS! Installer for 'Nexo_Test_1' finished with ExitCode 0 <<<
Triggering deployment binary sync...
Closing running Gratyfikant / Nexo processes to release file locks...
Syncing compiled plugin to all Nexo deployments in C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo...
Copying to deployment: KLINEX SP_Z_O_Oa4bfb87390e94eea9...
  >>> Copied successfully! Size: 247296 bytes | LastWrite: 11/08/2026 20:15:30
All deployment directories synchronized!
Target databases (Nexo_Test_1) installation and sync completed!
```

---

### 3.3 `sync_dll_to_all_deployments.ps1`

- **Location**: `scripts/sync_dll_to_all_deployments.ps1`
- **Purpose**: Directly synchronizes compiled assemblies and metadata into all local Gratyfikant Nexo Pro deployment cache folders without requiring a full database re-registration.

#### Script Parameters:
| Parameter | Type | Default Value | Description |
|---|---|---|---|
| `-KillRunningNexo` | `switch` | `false` | When enabled, forcefully terminates any running `Gratyfikant` or `Nexo` processes to release file locks on assemblies before copying. |

#### Synchronized Files:
- `HrAppka_Import_Pracowników.dll` (Core plugin assembly)
- `HrAppka_Import_Pracowników.pdb` (Debug symbols)
- `HrAppka_Import_Pracowników.plugin` (InsERT plugin declaration)
- `MiniExcel.dll` (Excel streaming ingestion dependency)

#### Target Directory Scanning:
The script scans `%LOCALAPPDATA%\InsERT\Deployments\Nexo\`:
```
C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo\
├── KLINEX SP_Z_O_Oa4bfb87390e94eea9\Binaries\
├── EUROSUPPORT_GROUPb3fe91820... \Binaries\
└── EUROSUPPORT OUTSOURCINGc741... \Binaries\
```
For every directory containing a `Binaries\` child folder, the script copies all four files using `[System.IO.File]::Copy(..., true)` to ensure atomic overwriting.

#### File Lock Handling:
If Gratyfikant Nexo is currently running and `-KillRunningNexo` was omitted, the Windows file system holds a mandatory lock on the DLL. The script catches the `IOException` and prints a clear diagnostic warning without crashing the entire script:
```text
WARNING: Could not overwrite DLL in KLINEX SP_Z_O_O... - process Nexo.exe is currently locking this file in memory! (The process cannot access the file because it is being used by another process.)
```

#### Invocation:
```powershell
powershell -ExecutionPolicy Bypass -File "scripts\sync_dll_to_all_deployments.ps1" -KillRunningNexo
```

---

### 3.4 `check_db_status.ps1`

- **Location**: `scripts/check_db_status.ps1`
- **Purpose**: Verifies that the plugin is properly registered in the deployment manifest of the target SQL Server database by executing a direct SQL query against the launcher schema.

#### Technical Mechanism:
Opens an ADO.NET connection to `100.127.240.115\INSERTGT` targeting `Nexo_Demo_1` (or test/production tenant) and executes:
```sql
SELECT DeploymentName, DeploymentManifest 
FROM InsLauncher.InstalledProducts
```
The script parses the returned XML string (`DeploymentManifest`), filters for occurrences of `"HrAppka"`, and outputs the exact identity and version registered in the database.

#### Invocation:
```powershell
powershell -ExecutionPolicy Bypass -File "scripts\check_db_status.ps1"
```

#### Expected Console Output:
```text
DeploymentName: Nexo
  -> Matches HrAppka:
     <Package Identity="HrAppka_Import_Pracowników-1.0.8.20" DisplayName="Import Pracowników" VersionDisplayName="1.0.8.20" />
```

---

### 3.5 `bump_version.py`

- **Location**: `scripts/bump_version.py`
- **Purpose**: Synchronizes version numbers across all configuration files to ensure build artifacts match the InsERT package manifest.

#### Target Files & Attributes:
1. `HrAppka Import Pracowników.csproj`:
   - `<Version>1.0.8.20</Version>`
   - `<AssemblyVersion>1.0.8.20</AssemblyVersion>`
   - `<FileVersion>1.0.8.20</FileVersion>`
2. `Konfiguracja.xml`:
   - `<Package Identity="HrAppka_Import_Pracowników-1.0.8.20" ... VersionDisplayName="1.0.8.20">`
   - `<PlikWynikowy>HrAppka_Import_Pracowników-1.0.8.20.exe</PlikWynikowy>`

#### Invocation:
```bash
python "C:\ NEXO SDK SFERA\HrAppka Import Pracowników\scripts\bump_version.py"
```

#### Expected Console Output:
```text
Updated HrAppka Import Pracowników.csproj -> 1.0.8.20: True
Updated Konfiguracja.xml -> 1.0.8.20: True
```

---

## 4. End-to-End Release & Deployment Workflows

### 4.1 Full Production Deployment Pipeline

When releasing a new version for operational use across all company databases, follow this standard sequence:

```powershell
# Step 1: Navigate to project root
cd "C:\ NEXO SDK SFERA\HrAppka Import Pracowników"

# Step 2: Increment version across .csproj and Konfiguracja.xml
python "scripts\bump_version.py"

# Step 3: Compile solution in Release mode
dotnet build "HrAppka Import Pracowników.csproj" -c Release --no-incremental

# Step 4: Run the complete test suite to ensure zero regressions
dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" -c Release

# Step 5: Build .mpkg deployment package and installer .exe
powershell -ExecutionPolicy Bypass -File "scripts\build_and_package_release.ps1"

# Step 6: Deploy to test database and verify
powershell -ExecutionPolicy Bypass -File "scripts\install_to_target_dbs.ps1" -Databases @("Nexo_Test_1")

# Step 7: Verify database registration manifest
powershell -ExecutionPolicy Bypass -File "scripts\check_db_status.ps1"

# Step 8: Deploy to production databases
powershell -ExecutionPolicy Bypass -File "scripts\install_to_target_dbs.ps1" -Databases @("Nexo_KLINEX SP_Z_O_O", "Nexo_EUROSUPPORT_GROUP", "Nexo_EUROSUPPORT OUTSOURCING")
```

---

### 4.2 Rapid Developer Hot-Reload (QuickSync)

During local feature development or UI styling adjustments, human developers and AI coding agents can bypass database manifest re-registration. The hot-reload loop completes in less than 3 seconds:

```powershell
# Step 1: Make C# or XAML code change
# Step 2: Compile Debug
dotnet build "HrAppka Import Pracowników.csproj" -c Debug

# Step 3: Push binaries to running Nexo deployment folders
powershell -ExecutionPolicy Bypass -File "scripts\sync_dll_to_all_deployments.ps1" -KillRunningNexo

# Step 4: Re-open Gratyfikant Nexo Pro
```

---

## 5. Script Parameter & Operational Matrix

| Script | Primary Responsibility | Input Parameters | Mandatory Pre-conditions | Produced Artifacts | Exit Behavior |
|---|---|---|---|---|---|
| `build_and_package_release.ps1` | Packages release archive and installer | None | Successful `dotnet build`; `PackageAndInstallerBuilder.exe` in SDK tools path | `Pakiet/*.mpkg`, `Instalator/*.exe` | `0` on success; throws exception on error |
| `install_to_target_dbs.ps1` | Registers package in SQL Server Launcher and updates manifests | `-Databases` (array), `-Server` (string), `-QuickSyncOnly` (switch) | Target SQL Server reachable; valid `.mpkg` in `Pakiet/` | Updated `InsLauncher.InstalledProducts` on target DBs; synced client files | `0` on success; logs warning on failure |
| `sync_dll_to_all_deployments.ps1` | Hot-syncs assemblies to local client cache | `-KillRunningNexo` (switch) | Built binaries present in `bin\Debug\` | Updated `.dll`, `.pdb`, `.plugin`, `MiniExcel.dll` in client folders | Non-zero warning on file lock; continues |
| `check_db_status.ps1` | Queries and outputs registered package info from SQL | None (parameterizable connection string) | SQL Server `100.127.240.115` reachable | Console text dump of deployment manifest | Catches and prints SQL exception |
| `bump_version.py` | Synchronizes version strings across manifests | Script parameters | Read/write permissions on `.csproj` and `.xml` | Updated version tags in project files | Returns `0`; prints status |

---

## 6. Troubleshooting & Diagnostic Commands

### 1. SQL Server Unreachable
```powershell
# Test TCP connectivity to SQL Server port 1433
Test-NetConnection -ComputerName 100.127.240.115 -Port 1433
```
- If connection fails: verify that VPN/Tailscale or LAN routing to `100.127.240.115` is active.

### 2. File Lock Preventing Binary Sync
```powershell
# Identify processes locking the DLL
Get-Process | Where-Object { $_.ProcessName -like "*Gratyfikant*" -or $_.ProcessName -like "*Nexo*" }
# Force termination
Stop-Process -Name Gratyfikant, Nexo -Force -ErrorAction SilentlyContinue
```

### 3. Verify Package Registration via SQL Query
```sql
SELECT 
    DeploymentName, 
    CAST(DeploymentManifest AS xml).query('/DeploymentManifest/AdditionalPackages/Package[@Identity="HrAppka_Import_Pracowników-1.0.8.20"]') AS PackageRegistration
FROM InsLauncher.InstalledProducts
WHERE DeploymentName = 'Nexo';
```

---
*Document author: Antigravity Technical Documentation Worker (`worker_doc_m2`)*  
*Last verified against codebase: November 2026*
