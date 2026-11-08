# Legacy & Redundant Code Audit

**Solution**: HrAppka Import Pracowników (v1.0.8.20)  
**Target Framework**: `.NET 8.0` (`net8.0-windows10.0.17763.0`)  
**SDK Dependency**: InsERT Gratyfikant Nexo Pro SDK 62.0.0.9537  
**Host Application**: InsERT Gratyfikant Nexo Pro (`Gratyfikant nexo.exe`)  
**Scope**: Comprehensive Technical Audit of Deprecated, Dead, and Redundant Code

---

## Table of Contents
1. [Executive Summary](#1-executive-summary)
2. [Architectural Comparison: Legacy v1 vs Modern v2](#2-architectural-comparison-legacy-v1-vs-modern-v2)
   - [2.1 Side-by-Side Comparison Matrix](#21-side-by-side-comparison-matrix)
   - [2.2 Inactivation Proof in GrupaOperacji.cs](#22-inactivation-proof-in-grupaoperacjics)
   - [2.3 Deep-Dive into v1 Architectural Flaws](#23-deep-dive-into-v1-architectural-flaws)
3. [Itemized Catalog of Obsolete & Redundant Components](#3-itemized-catalog-of-obsolete--redundant-components)
   - [3.1 OperacjaImportuPracownikow.cs (v1 Engine)](#31-operacjaimportupracownikowcs-v1-engine)
   - [3.2 Program.cs Diagnostic Methods & CLI Harness](#32-programcs-diagnostic-methods--cli-harness)
   - [3.3 Obsolete Maintenance Scripts](#33-obsolete-maintenance-scripts)
   - [3.4 Root ParametryInstalacji.txt Configuration](#34-root-parametryinstalacjitxt-configuration)
   - [3.5 Legacy csproj & XML Manifest Elements](#35-legacy-csproj--xml-manifest-elements)
   - [3.6 In-File Obsolete Dialogs in OperacjaImportuPracownikow20.cs](#36-in-file-obsolete-dialogs-in-operacjaimportupracownikow20cs)
4. [Safe 4-Phase Step-by-Step Pruning Strategy](#4-safe-4-phase-step-by-step-pruning-strategy)
   - [Phase 1: Standalone Script & Batch Cleanup](#phase-1-standalone-script--batch-cleanup)
   - [Phase 2: v1 Monolithic Engine Pruning](#phase-2-v1-monolithic-engine-pruning)
   - [Phase 3: Program.cs CLI Decoupling & Library Mode Transition](#phase-3-programcs-cli-decoupling--library-mode-transition)
   - [Phase 4: In-File Modernization of OperacjaImportuPracownikow20.cs](#phase-4-in-file-modernization-of-operacjaimportupracownikow20cs)
5. [Refactoring Risk Assessment & Verification Protocol](#5-refactoring-risk-assessment--verification-protocol)

---

## 1. Executive Summary

During the evolution of the **HrAppka Import Pracowników** plugin from early prototyping to an enterprise-grade production solution (v1.0.8.20), substantial architectural refactoring occurred. The original v1 implementation (`OperacjaImportuPracownikow.cs`) was entirely superseded by the modern Sfera Native 2.0 engine (`OperacjaImportuPracownikow20.cs`), supported by modular domain services (`ContractCreationService`, `SferaDictionaryCache`, `UniversalFileLoader`, `ZusDeclarationBatchService`).

However, legacy source files, diagnostic test harnesses, obsolete deployment batch scripts, and hardcoded development artifacts remain in the repository. While these artifacts do not compromise runtime execution (because Gratyfikant Nexo only invokes the registered v2 engine), they increase repository bloat, confuse onboarding developers, and introduce potential confusion regarding system behavior.

This audit provides:
1. An exhaustive architectural analysis comparing v1 and v2.
2. Verified proof of v1 disconnection from runtime entrypoints.
3. An explicit, itemized catalog of all dead and deprecated components.
4. A safe, 4-phase incremental pruning roadmap for future cleanup.

---

## 2. Architectural Comparison: Legacy v1 vs Modern v2

### 2.1 Side-by-Side Comparison Matrix

| Architectural Dimension | Legacy v1 (`OperacjaImportuPracownikow.cs`) | Modern v2 (`OperacjaImportuPracownikow20.cs`) | Architectural Impact / Benefit in v2 |
|---|---|---|---|
| **Source Line Count** | **1,570 lines** | **3,156 lines** (plus modular services) | v2 provides comprehensive validation, interactive UI, pre-caching, and modular service separation. |
| **Plugin Registration** | **Completely Unregistered** (0 registrations in `GrupaOperacji.cs`). | **Active Runtime Entry Point** registered in `GrupaOperacji.cs` line 20. | v1 is dead code in production Gratyfikant Nexo Pro environments. |
| **Operation Name** | `"Import Pracowników (v1 Legacy)"` (line 24). | `"Import Pracowników"` (line 37). | Clear branding aligned with production UI specifications. |
| **Contract Creation** | Monolithic procedural logic inlined in `ZaimportujRow` (lines 906–1070). | Delegated to `ContractCreationService.Instance.CreateContract(...)` (lines 1963–1970). | Decouples contract business rules; encapsulates salary calculations and template resolution; testable. |
| **Contract Templates** | Hardcoded RTF default template `uchwyt.TresciWydrukuRTF().DaneDomyslne.UCPZlecenieStale` (line 924). | Dynamic enumeration via `ISzablonyUmowPracowniczychGr`, template lookup, and operator override in UI. | Supports custom user-defined contract templates and company-specific civil contracts. |
| **Entity Framework Context** | Queried entities directly from `IUchwyt` managers, triggering cross-context detached entity exceptions (`IEntityChangeTracker`). | Resolves BO's internal `ObjectContext` (`ModelDanychContainer`) and re-queries entities in the same UoW. | Eliminates fatal `InvalidOperationException` collisions upon `Zapisz()`. |
| **User Interface & Concurrency** | Synchronous modal dialog (`IOknoOperacjiZbiorczej`) blocking the UI thread (lines 304–311). | Asynchronous modal WPF window (`ImportPreviewWindow`) with background worker `Task.Run` and real-time progress reporting. | Non-blocking execution; responsive UI; operator review of validation warnings; cancellation support. |
| **File Ingestion & Parsing** | Excel-only (`MiniExcel.Query`); OpenFileDialog filter strictly `*.xlsx` (lines 70, 87). | Universal ingestion via `UniversalFileLoader.LoadFile(sciezka)` supporting `.xlsx` and multiline RFC 4180 `.csv`. | Handles client CSV files with multiline street addresses, quoted semicolons, and UTF-8 BOM. |
| **Database Performance** | N+1 queries: queries `WszystkieOsoby()`, `Panstwa()`, `Obywatelstwa()`, `Waluty()`, `IFlagiWlasne` per row. | In-memory pre-caching via `SferaDictionaryCache.Instance.Initialize(uchwyt)` and O(1) hash maps. | Reduces database roundtrips by up to 95% on large employee datasets. |
| **ZUS Declaration Generation** | Immediate per-row save: opens `IDeklaracjaZUS` and executes `deklBO.Zapisz()` inside the per-employee loop (lines 1076–1185). | Single-batch consolidation: accumulates requests and executes single transaction via `ZusDeclarationBatchService`. | Eliminates SQL Server lock contention; prevents fragmented declaration batches. |
| **Bank Account Handling** | Attempted to query/create `Firma` for the bank podmiot (lines 807–824), risking corporate table duplication. | Sets `PodmiotBankowy = null` and stores bank name in `Opis`/`Nazwa`, strictly populating IBAN on personal account. | Prevents database clutter and external firm table contamination. |
| **SDK Paths & Reflection** | Contains hardcoded legacy reflection path: `Assembly.LoadFrom(@"C:\ NEXO SDK SFERA\nexoSDK_60.1.1.9292\Bin\InsERT.Moria.API.Private.dll")` (line 1136). | Strongly typed references configured in `.csproj` (`nexoSdkBinPath`). | Eliminates hardcoded local paths to obsolete SDK versions. |

---

### 2.2 Inactivation Proof in `GrupaOperacji.cs`

Direct inspection of `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\GrupaOperacji.cs` proves that legacy v1 (`OperacjaImportuPracownikow`) is completely disconnected from Gratyfikant Nexo:

```csharp
// C:\ NEXO SDK SFERA\HrAppka Import Pracowników\GrupaOperacji.cs
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Rozszerzanie.Operacje;
using System;
using System.Collections.Generic;

namespace HrAppka_Import_Pracowników
{
    public class GrupaOperacji : IGrupaOperacji
    {
        public IDostawcaPluginow Dostawca => new DostawcaPluginow();

        public Guid Identyfikator => new Guid("8e2be35b-12d8-4f81-9b16-562a93bbef88");

        public string Nazwa => "Import Pracowników";

        public string Opis => "Operacje importu pracowników do Gratyfikant Nexo Pro";

        public IEnumerable<Operacja> Operacje { get; } = new Operacja[]
        {
            new OperacjaImportuPracownikow20() // LINE 20: ONLY v2 IS REGISTERED!
        };
    }
}
```

Because `OperacjaImportuPracownikow` is not instantiated in `Operacje`, Gratyfikant Nexo's plugin discovery loader never exposes or invokes it. The entire 1,570 lines of `OperacjaImportuPracownikow.cs` constitute completely dead code.

---

### 2.3 Deep-Dive into v1 Architectural Flaws

#### 1. Inlined Monolithic Contract Creation (Lines 906–1070)
In v1, lines 906–1070 contained over 160 lines of procedural contract setup inside the row processing loop:
```csharp
// Legacy v1 inlined creation (OperacjaImportuPracownikow.cs:906)
using (var umowa = uchwyt.UmowyPracowniczeGr().Utworz(umowaZlecenie))
{
    var baseUmowa = umowa.Dane;
    baseUmowa.Pracownik = employeePodmiot.Osoba.Pracownik;
    baseUmowa.RodzajUmowyPracowniczej = umowaZlecenie.DozwoloneRodzajeUmowyPracowniczej.FirstOrDefault(x => x.Nazwa == "Na czas określony");
    baseUmowa.WzorzecTresciUmowyPracowniczej = uchwyt.TresciWydrukuRTF().DaneDomyslne.UCPZlecenieStale;
    // ... manual flag configuration ...
    if (!umowa.Zapisz()) { ... }
}
```
**Why it failed**: Modifying contract terms, template IDs, or wage calculation algorithms required editing the monolithic import file. It lacked template enumeration and could not support company-specific custom contract definitions.

#### 2. Cross-DbContext Detached Entity Collisions
When v1 assigned `employeePodmiot.Osoba.Pracownik` (queried from `uchwyt.Podmioty()`) directly to `baseUmowa.Pracownik` (in `uchwyt.UmowyPracowniczeGr()`), Entity Framework threw:
> `InvalidOperationException: An entity object cannot be referenced by multiple instances of IEntityChangeTracker.`

**Modern Solution in v2**: `ContractCreationService` extracts the BO's internal `ObjectContext` (`ModelDanychContainer`) via reflection and re-queries the `Pracownik` within that exact DataDomain:
```csharp
// Modern v2 Context Attachment (Services/ContractCreationService.cs:283-315)
var prop = u.GetType().GetProperty("ObjectContext", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
var context = prop?.GetValue(u) as ModelDanychContainer;

if (context != null)
{
    int pracownikId = employeePodmiot.Osoba.Pracownik.Id;
    var attachedPracownik = context.Pracownicy.FirstOrDefault(p => p.Id == pracownikId);
    u.Pracownik = attachedPracownik ?? employeePodmiot.Osoba.Pracownik;
}
```

#### 3. Immediate Per-Row ZUS Declaration Saves
In v1 (lines 1076–1185), every processed row opened, modified, and committed a separate ZUS declaration:
```csharp
// Legacy v1 immediate save per row
var existingDecl = dekZusMgr.Dane.Wszystkie().Where(...).FirstOrDefault();
IDeklaracjaZUS deklBO = existingDecl != null ? dekZusMgr.Znajdz(existingDecl) : dekZusMgr.UtworzDeklaracjeZgloszeniowa();
using (deklBO)
{
    // ... add 1 record ...
    deklBO.Zapisz(); // Executed for EVERY employee row!
}
```
**Why it failed**: Importing 300 employees triggered 300 separate write transactions against the database, resulting in transaction timeouts, SQL locking, and fragmented registration documents.

---

## 3. Itemized Catalog of Obsolete & Redundant Components

The following catalog identifies all deprecated components, obsolete helper classes, and dead scripts currently present in the repository:

```
+---------------------------------------------------------------------------------------------------------------------+
|                                 Itemized Obsolete & Redundant Code Catalog                                          |
+----+------------------------------------+--------------------------+-----------------------+------------------------+
| #  | Component / File                   | Category                 | Current Status        | Replacement in v2      |
+----+------------------------------------+--------------------------+-----------------------+------------------------+
| 1  | OperacjaImportuPracownikow.cs      | Monolithic Import Engine | Dead Code (1570 lines)| OperacjaImportuPracownikow20 |
| 2  | Program.cs (AutoInstaller)         | Obsolete CLI Installer   | Deprecated Harness    | install_to_target_dbs.ps1 |
| 3  | Program.cs (SferaRunner)           | Diagnostic Console CLI   | Deprecated Harness    | In-tree xUnit Tests    |
| 4  | Instaluj_Automatycznie.bat         | Obsolete Batch Script    | Dead Script (v1.0.0.2)| build_and_package_release.ps1 |
| 5  | install.ps1                        | Obsolete Script          | Dead Script           | install_to_target_dbs.ps1 |
| 6  | update_methods.ps1                 | Scratch Maintenance Tool | Dead Artifact         | In-tree C# methods     |
| 7  | ParametryInstalacji.txt (Root)     | Static Config File       | Redundant Artifact    | Generated dynamically  |
| 8  | WpfWarningResolverDialog           | Modal Popup Helper       | Redundant UI          | ImportPreviewWindow    |
+----+------------------------------------+--------------------------+-----------------------+------------------------+
```

---

### 3.1 `OperacjaImportuPracownikow.cs` (v1 Engine)

- **File Location**: `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\OperacjaImportuPracownikow.cs`
- **Line Count**: 1,570 lines.
- **Current State**: Dead code. Not registered in `GrupaOperacji.cs`. Never instantiated at runtime.
- **Reason for Obsolescence**: Completely replaced by `OperacjaImportuPracownikow20.cs`. Contains obsolete SDK paths (`nexoSDK_60.1.1.9292\Bin\InsERT.Moria.API.Private.dll`), synchronous UI blocking, and lack of CSV support.
- **Replacement in v2**: `OperacjaImportuPracownikow20.cs`, `ContractCreationService.cs`, `UniversalFileLoader.cs`, `ZusDeclarationBatchService.cs`.
- **Pruning Risk**: **Low Risk**. Can be deleted without impacting Gratyfikant runtime or test suite.

---

### 3.2 `Program.cs` Diagnostic Methods & CLI Harness

- **File Location**: `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\Program.cs`
- **Line Count**: 607 lines.
- **Classes**: `Program` (lines 16–120), `AutoInstaller` (lines 122–403), `SferaRunner` (lines 405–607).
- **Current State**: Deprecated CLI entry point used during initial development.
- **Reason for Obsolescence**:
  1. `AutoInstaller`: Implements `--install` and `--uninstall` using raw ADO.NET and Mox Launcher APIs. While currently invoked by `scripts/install_to_target_dbs.ps1`, it forces the project to compile as `<OutputType>Exe</OutputType>` rather than a standard `<OutputType>Library</OutputType>` class library.
  2. `SferaRunner`: Contains hardcoded personal paths (e.g. `line 411: @"C:\Users\zasc2\Downloads\Telegram Desktop\Eksport - Umowy - ANDROS.xlsx"`) and reflection dumps of Moria interfaces.
- **Replacement in v2**: Gratyfikant runtime plugin discovery (`DostawcaPluginow.cs`), automated PowerShell scripts (`scripts/install_to_target_dbs.ps1`), and automated xUnit tests in `tests/`.
- **Pruning Risk**: **Medium Risk**. `install_to_target_dbs.ps1` currently relies on `dotnet HrAppka_Import_Pracowników.dll --install`. Before pruning `AutoInstaller`, its logic should be converted to a standalone CLI tool or pure PowerShell module.

---

### 3.3 Obsolete Maintenance Scripts

#### A. `Instaluj_Automatycznie.bat`
- **Location**: `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\Instaluj_Automatycznie.bat`
- **Content**: 24 lines.
- **Reason for Obsolescence**: Hardcodes obsolete plugin version `1.0.0.2` (line 22: `"Plugin version 1.0.0.2 has been successfully installed and registered!"`). Runs unparameterized `dotnet run --project ... -- --install`.
- **Replacement**: `scripts/build_and_package_release.ps1` and `scripts/install_to_target_dbs.ps1`.
- **Pruning Risk**: **Zero Risk**.

#### B. `install.ps1`
- **Location**: `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\install.ps1`
- **Content**: 13 lines.
- **Reason for Obsolescence**: Hardcodes array of 3 databases (`Nexo_KLINEX SP_Z_O_O`, `Nexo_EUROSUPPORT_GROUP`, `Nexo_EUROSUPPORT OUTSOURCING`). Lacks packaging verification, parameter support, error handling, or process termination.
- **Replacement**: `scripts/install_to_target_dbs.ps1`.
- **Pruning Risk**: **Zero Risk**.

#### C. `update_methods.ps1`
- **Location**: `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\update_methods.ps1`
- **Content**: 118 lines.
- **Reason for Obsolescence**: One-off development scratch script that used regular expressions to patch `PreAnalyzeRows` into `OperacjaImportuPracownikow20.cs`.
- **Replacement**: Already permanently integrated into `OperacjaImportuPracownikow20.cs`.
- **Pruning Risk**: **Zero Risk**.

---

### 3.4 Root `ParametryInstalacji.txt` Configuration

- **Location**: `C:\ NEXO SDK SFERA\HrAppka Import Pracowników\ParametryInstalacji.txt`
- **Reason for Obsolescence**: Static INI configuration file containing plaintext SQL credentials. `scripts/install_to_target_dbs.ps1` dynamically generates and overwrites this file for each targeted database.
- **Replacement**: Parameterized inputs in `scripts/install_to_target_dbs.ps1`.
- **Pruning Risk**: **Low Risk**. Can be removed from root once scripts generate it directly into temporary directories or `bin\Debug\`.

---

### 3.5 Legacy csproj & XML Manifest Elements

1. **`HrAppka Import Pracowników.csproj`**:
   - `<OutputType>Exe</OutputType>`: Enforced only by `Program.cs`. Pure Sfera plugins should be `<OutputType>Library</OutputType>`.
   - `<nexoSdkNarzedziaPath>C:\ NEXO SDK SFERA\nexoSDK_61.0.4.9400\Narzedzia\</nexoSdkNarzedziaPath>`: References version 61.0.4, whereas the main SDK path is 62.0.0.9537.
2. **`Konfiguracja.xml`**:
   - Contains legacy element `<ParametryInstalacji StworzPlikiBat="Tak" AutomatycznaInstalacja="Nie">` referencing `ParametryInstalacji.txt`.

---

### 3.6 In-File Obsolete Dialogs in `OperacjaImportuPracownikow20.cs`

- **Class**: `WpfWarningResolverDialog` (lines 1471–1579 in `OperacjaImportuPracownikow20.cs`).
- **Reason for Obsolescence**: Originally designed to display synchronous modal message boxes (`System.Windows.Application.Current.Dispatcher.Invoke`) asking the operator about birthdate discrepancies, gender mismatches, or missing tax offices per row.
- **Replacement**: The interactive WPF preview DataGrid (`UI/ImportPreviewWindow.xaml` and `ImportPreviewViewModel.cs`) presents all warnings inline and allows the operator to resolve conflicts directly in the table before initiating the import.
- **Pruning Risk**: **Low Risk**.

---

## 4. Safe 4-Phase Step-by-Step Pruning Strategy

To eliminate technical debt without disrupting active development or breaking the release pipeline, the following 4-phase sequence is recommended:

```
┌────────────────────────────────────────────────────────────────────────┐
│ Phase 1: Standalone Script & Batch Cleanup                             │
│ • Delete: `update_methods.ps1`, `install.ps1`, `Instaluj_Automatycznie.bat` │
│ • Verification: Execute `scripts/build_and_package_release.ps1`        │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
┌───────────────────────────────────▼────────────────────────────────────┐
│ Phase 2: v1 Monolithic Engine Pruning                                  │
│ • Confirm `GrupaOperacji.cs` only registers `OperacjaImportuPracownikow20` │
│ • Delete `OperacjaImportuPracownikow.cs` (1,570 lines)                 │
│ • Verification: `dotnet build -c Debug`, `dotnet build -c Release`     │
│ • Execute `dotnet test` to confirm test suite unaffected               │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
┌───────────────────────────────────▼────────────────────────────────────┐
│ Phase 3: Program.cs CLI Decoupling & Library Mode Transition          │
│ • Move `AutoInstaller` logic to pure PowerShell / separate utility     │
│ • Change `<OutputType>Exe</OutputType>` to `<OutputType>Library</OutputType>` │
│ • Move `SferaRunner` to standalone diagnostic CLI project              │
│ • Verification: Run `scripts/install_to_target_dbs.ps1` pipeline       │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
┌───────────────────────────────────▼────────────────────────────────────┐
│ Phase 4: In-File Modernization of OperacjaImportuPracownikow20.cs    │
│ • Remove obsolete `WpfWarningResolverDialog` (lines 1471–1579)         │
│ • Clean up unused synchronous helper methods                           │
│ • Align SDK tools path in `.csproj` to 62.0.0.9537                     │
│ • Verification: Execute all Tier 1–4 tests and test import run         │
└────────────────────────────────────────────────────────────────────────┘
```

---

### Phase 1: Standalone Script & Batch Cleanup
1. **Action**: Remove `update_methods.ps1`, `install.ps1`, and `Instaluj_Automatycznie.bat` from the project root.
2. **Verification Command**:
   ```powershell
   # Confirm scripts build pipeline continues to function
   powershell -ExecutionPolicy Bypass -File "scripts\build_and_package_release.ps1"
   ```
3. **Safety Guarantee**: None of these files are referenced by `.csproj`, `Konfiguracja.xml`, or the test suite.

---

### Phase 2: v1 Monolithic Engine Pruning
1. **Action**: Remove `OperacjaImportuPracownikow.cs` (1,570 lines).
2. **Pre-check**: Verify `GrupaOperacji.cs` contains zero references to `OperacjaImportuPracownikow`.
3. **Verification Command**:
   ```powershell
   dotnet build "HrAppka Import Pracowników.csproj" -c Debug --no-incremental
   dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj"
   ```
4. **Safety Guarantee**: Because `GrupaOperacji.cs` exclusively instantiates `OperacjaImportuPracownikow20`, deleting v1 does not alter runtime behavior or compiled output.

---

### Phase 3: Program.cs CLI Decoupling & Library Mode Transition
1. **Action**:
   - Extract `AutoInstaller.Install()` and `AutoInstaller.Uninstall()` into a PowerShell script or separate administrative utility project.
   - Update `scripts/install_to_target_dbs.ps1` to invoke the new utility.
   - In `HrAppka Import Pracowników.csproj`, change:
     ```xml
     <OutputType>Library</OutputType>
     ```
   - Remove or relocate `Program.cs`.
2. **Verification Command**:
   ```powershell
   dotnet build "HrAppka Import Pracowników.csproj" -c Release
   powershell -ExecutionPolicy Bypass -File "scripts\build_and_package_release.ps1"
   powershell -ExecutionPolicy Bypass -File "scripts\install_to_target_dbs.ps1" -Databases @("Nexo_Test_1")
   ```

---

### Phase 4: In-File Modernization of `OperacjaImportuPracownikow20.cs`
1. **Action**:
   - Remove nested class `WpfWarningResolverDialog` (lines 1471–1579).
   - Ensure all user prompt fallbacks redirect through `ImportPreviewViewModel`.
   - Update `nexoSdkNarzedziaPath` in `.csproj` from `61.0.4.9400` to `62.0.0.9537`.
2. **Verification Command**:
   ```powershell
   dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" --logger "console;verbosity=normal"
   ```

---

## 5. Refactoring Risk Assessment & Verification Protocol

| Refactoring Step | Target Component | Potential Risk | Prevention & Verification Protocol |
|---|---|---|---|
| **Phase 1** | Standalone scripts | Low / None | Verify git status confirms no referenced files removed. |
| **Phase 2** | `OperacjaImportuPracownikow.cs` | Low | Verify clean compilation of `HrAppka Import Pracowników.csproj` and pass of all xUnit tests. |
| **Phase 3** | `Program.cs` & OutputType | Medium | Test `install_to_target_dbs.ps1` on `Nexo_Test_1` staging database before applying to production. |
| **Phase 4** | In-file dialog removal | Low | Verify WPF preview window operates correctly under interactive manual testing. |

---
*Document author: Antigravity Technical Documentation Worker (`worker_doc_m2`)*  
*Last verified against codebase: November 2026*
