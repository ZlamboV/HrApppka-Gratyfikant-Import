# HrAppka Import Pracowników - Gratyfikant Nexo Pro Sfera Plugin

[![Version](https://img.shields.io/badge/version-1.0.8.20-blue.svg)](HrAppka%20Import%20Pracowników.csproj)
[![Framework](https://img.shields.io/badge/framework-.NET%208.0%20(Windows%2010.0.17763.0)-purple.svg)](HrAppka%20Import%20Pracowników.csproj)
[![SDK](https://img.shields.io/badge/InsERT%20Nexo%20SDK-62.0.0.9537-orange.svg)](docs/architecture.md)
[![Database](https://img.shields.io/badge/SQL%20Server-100.127.240.115%5CINSERTGT-green.svg)](docs/deployment_and_scripts.md)

An enterprise-grade in-process Sfera plugin for **InsERT Gratyfikant Nexo Pro**, providing end-to-end ingestion, validation, normalization, and automated import of external workforce recruitment spreadsheets (HrAppka, WorkService, custom Excel/CSV files) into Gratyfikant Nexo's relational HR and payroll domain model.

---

## Table of Contents
1. [Executive Summary](#1-executive-summary)
2. [Master Documentation Index](#2-master-documentation-index)
3. [Quick Start Guide](#3-quick-start-guide)
4. [System Architecture Overview](#4-system-architecture-overview)
5. [Core Functional Modules](#5-core-functional-modules)
6. [Target Database Topology](#6-target-database-topology)
7. [Deployment & Automation Scripts](#7-deployment--automation-scripts)
8. [Maintenance & Git Workflow](#8-maintenance--git-workflow)

---

## 1. Executive Summary

The **HrAppka Import Pracowników** plugin automates the complex, error-prone workflow of onboarding personnel from external staffing portals into **InsERT Gratyfikant Nexo Pro**. Built directly on the native InsERT Mox/Sfera SDK architecture, the plugin operates in-process inside `Gratyfikant nexo.exe`.

### Key Capabilities:
- **Intelligent Identity Resolution**: Dual-key duplicate detection matching against existing personnel via 11-digit PESEL or foreign identity document / passport number.
- **Contract Lifecycle Automation**: Dynamic instantiation of civil contracts (`Umowa zlecenie`) with template resolution, variable wage definitions, student under-26 tax exemptions, and custom contract duration bounds.
- **ZUS Registration Batching**: Single-transaction consolidation of ZUS registration declarations (`ZUA` / `ZZA`) without database lock contention.
- **Resilient Data Sanitization**: Universal ingestion supporting Excel (`.xlsx`) and RFC 4180 multiline CSV, paired with compiled regex address splitting, SQLite postal code enrichment (`DaneAdresowe.db`), and Modulo 97 IBAN checksum validation.
- **Interactive Operator Oversight**: Virtualized WPF modal interface (`ImportPreviewWindow`) with color-coded status badges, real-time warning diagnostics, inline template/department selection, and safe background execution via `Task.Run`.

---

## 2. Master Documentation Index

Comprehensive technical documentation is maintained in the [`docs/`](docs/) directory:

```
docs/
├── architecture.md                <- Global system architecture, Sfera 2.0 integration & Mox DataDomain
├── data_flows.md                  <- Visual Mermaid & ASCII diagrams of import lifecycle & validation gates
├── deployment_and_scripts.md      <- Full documentation of all 5 automation scripts & SQL deployment topology
├── onboarding_dev_ai.md           <- End-to-end developer & AI onboarding guide + 7-scenario troubleshooting matrix
├── legacy_code_audit.md           <- Legacy audit v1 vs v2, itemized catalog & 4-phase pruning strategy
└── modules/
    ├── sfera_import.md            <- Sfera Native 2.0 import engine, employee creation & duplicate resolution
    ├── contracts.md               <- Contract Creation Service, entity attachment & civil contract rules
    ├── normalization.md           <- File loaders, RFC 4180 CSV, address normalizer, SQLite & IBAN resolution
    └── ui_preview.md              <- Interactive WPF UI, DataGrid virtualization, MVVM ViewModel & converters
```

### Direct Documentation Links:
- 🏛️ **[System Architecture Guide](docs/architecture.md)**: Deep dive into the InsERT Mox Unit of Work lifecycle, Entity Framework multi-context sponsorship, STA threading, and N+1 query elimination.
- 🔄 **[End-to-End Data Flows](docs/data_flows.md)**: Visual sequence and flow diagrams detailing file ingestion, validation gates, duplicate conflict resolution, contract creation, and ZUS declaration commits.
- 🚀 **[Deployment & Automation Scripts Guide](docs/deployment_and_scripts.md)**: Comprehensive manual for `build_and_package_release.ps1`, `install_to_target_dbs.ps1`, `sync_dll_to_all_deployments.ps1`, `check_db_status.ps1`, and `bump_version.py`.
- 🧑‍💻 **[Developer & AI Agent Onboarding Guide](docs/onboarding_dev_ai.md)**: Environment setup (.NET 8, SDK 62), build commands, xUnit test tiers 1–4, and a systematic 7-scenario troubleshooting matrix.
- 🧹 **[Legacy Code Audit & Pruning Strategy](docs/legacy_code_audit.md)**: Side-by-side comparison of deprecated v1 (`OperacjaImportuPracownikow.cs`) vs modern v2 (`OperacjaImportuPracownikow20.cs`), catalog of dead scripts, and 4-phase pruning roadmap.
- ⚙️ **[Sfera 2.0 Import Module Guide](docs/modules/sfera_import.md)**: Technical breakdown of `OperacjaImportuPracownikow20.cs`, employee identity lifecycle, address/bank/NFZ mapping, and error diagnostics.
- 📝 **[Contract Creation Service Guide](docs/modules/contracts.md)**: In-depth reference for `ContractCreationService.cs`, `ISzablonyUmowPracowniczychGr`, BO DataDomain entity attachment, and civil contract calculations.
- 🔍 **[Data Normalization & Ingestion Guide](docs/modules/normalization.md)**: Complete guide to `UniversalFileLoader.cs`, `Rfc4180CsvReader.cs`, regex address parsing, SQLite postal code enrichment, and IBAN validation.
- 🖥️ **[Interactive UI & Preview Window Guide](docs/modules/ui_preview.md)**: Architecture of `ImportPreviewWindow.xaml`, MVVM bindings, DataGrid UI virtualization, Win32 modal hosting, and progress reporting.

---

## 3. Quick Start Guide

### Prerequisites
- Windows 10/11 x64
- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) (`net8.0-windows10.0.17763.0`)
- InsERT Gratyfikant Nexo Pro SDK 62.0.0.9537
- Network connectivity to MS SQL Server `100.127.240.115\INSERTGT`

### 1. Build Debug
```powershell
cd "C:\ NEXO SDK SFERA\HrAppka Import Pracowników"
dotnet build "HrAppka Import Pracowników.csproj" -c Debug
```

### 2. Run Test Suite
```powershell
dotnet test "C:\ NEXO SDK SFERA\tests\ImportPracownikow.Tests.csproj" --logger "console;verbosity=normal"
```

### 3. Rapid Developer Hot-Reload (QuickSync)
Push freshly compiled assemblies directly into running Gratyfikant Nexo deployment folders:
```powershell
powershell -ExecutionPolicy Bypass -File "scripts\sync_dll_to_all_deployments.ps1" -KillRunningNexo
```

### 4. Build Full Release Package & Deploy
```powershell
# 1. Package .mpkg release package and standalone .exe installer
powershell -ExecutionPolicy Bypass -File "scripts\build_and_package_release.ps1"

# 2. Deploy to staging/test database (Nexo_Test_1)
powershell -ExecutionPolicy Bypass -File "scripts\install_to_target_dbs.ps1" -Databases @("Nexo_Test_1")

# 3. Verify deployment manifest in SQL Server
powershell -ExecutionPolicy Bypass -File "scripts\check_db_status.ps1"
```

---

## 4. System Architecture Overview

The solution adheres to a modular, layered architecture separating user interface presentation, file ingestion, business normalization, and InsERT Mox Sfera domain operations:

```
+-------------------------------------------------------------------------------------------------------+
|                                    System Architecture Topology                                       |
+-------------------------------------------------------------------------------------------------------+

                                  [User Spreadsheet / CSV]
                                             │
                                             ▼
                       +───────────────────────────────────────────+
                       |       Ingestion & Normalization           |
                       |  - UniversalFileLoader / RFC 4180 CSV     |
                       |  - AddressNormalizer (Regex Splitting)    |
                       |  - AddressEnrichmentService (SQLite)      |
                       |  - BankClearingResolver (IBAN Modulo 97)  |
                       +─────────────────────┬─────────────────────+
                                             │
                                             ▼
                       +───────────────────────────────────────────+
                       |        Interactive WPF UI Preview         |
                       |  - ImportPreviewWindow (Virtualized Grid) |
                       |  - ImportPreviewViewModel (MVVM Bindings) |
                       |  - Win32 Modal Window Hosting (HWND Owner)|
                       |  - Real-time Validation & Operator Review |
                       +─────────────────────┬─────────────────────+
                                             │ Operator Clicks "Importuj"
                                             ▼ (Task.Run Background Thread)
                       +───────────────────────────────────────────+
                       |        Sfera Native 2.0 Engine            |
                       |  - OperacjaImportuPracownikow20           |
                       |  - SferaDictionaryCache (N+1 Elimination) |
                       |  - Employee Identity Deduplication        |
                       +───────────────┬───────────────────┬───────+
                                       │                   │
                        Attach Context │                   │ Batch Declarations
                                       ▼                   ▼
     +───────────────────────────────────+       +───────────────────────────────────+
     |      ContractCreationService      |       |    ZusDeclarationBatchService     |
     |  - EF ModelDanychContainer Link   |       |  - Consolidated ZUA/ZZA Creation  |
     |  - Civil Contract Lifecycle       |       |  - Single-Transaction Commit      |
     |  - Template & Remuneration Rules  |       +─────────────────┬─────────────────+
     +─────────────────┬─────────────────+                         │
                       │                                           │
                       └───────────────────┬───────────────────────┘
                                           │
                                           ▼
                       +───────────────────────────────────────────+
                       |        InsERT Gratyfikant Nexo Pro        |
                       |          Relational SQL Database          |
                       +-------------------------------------------+
```

---

## 5. Core Functional Modules

| Module | Primary Implementation | Key Architectural Responsibility |
|---|---|---|
| **Sfera Native 2.0 Engine** | `OperacjaImportuPracownikow20.cs` | Coordinates the full import workflow, performs duplicate matching (PESEL/Passport), drives employee creation/updates, and reports progress. |
| **Contract Creation Service** | `Services/ContractCreationService.cs` | Creates civil contracts, resolves templates via `ISzablonyUmowPracowniczychGr`, and attaches foreign entities within the BO's `ModelDanychContainer`. |
| **Dictionary Cache** | `Services/SferaDictionaryCache.cs` | Pre-caches countries, citizenships, tax offices, currencies, and departments, eliminating N+1 database roundtrips. |
| **ZUS Declaration Batching** | `Services/ZusDeclarationBatchService.cs` | Accumulates employee registration requests and commits all ZUS ZUA/ZZA declarations in a single atomic transaction. |
| **Data Ingestion** | `Services/UniversalFileLoader.cs` | Auto-detects file types, delegating `.xlsx` to MiniExcel and `.csv` to the multiline RFC 4180 stream reader. |
| **Address Normalization** | `Services/AddressNormalizer.cs` | Uses compiled regular expressions to separate combined address lines into Street, House Number, and Apartment Number. |
| **Postal Code Enrichment** | `Services/AddressEnrichmentService.cs` | Queries an offline SQLite database (`DaneAdresowe.db`) to auto-populate missing province, county, and municipality from postal code. |
| **Bank Account Validation** | `Services/BankClearingResolver.cs` | Validates Polish and international IBANs using Modulo 97 arithmetic and maps bank clearing numbers to institution names. |
| **Interactive WPF UI** | `UI/ImportPreviewWindow.xaml` | Virtualized DataGrid presentation with MVVM bindings, inline status badges, template overrides, and background execution. |
| **Modal Hosting Service** | `UI/Services/ModalWindowHostingService.cs` | Ensures WPF modal dialogs attach to Gratyfikant's main window handle via Win32 HWND fallback and enforce STA threading. |

---

## 6. Target Database Topology

The solution connects to a centralized SQL Server hosting both the InsERT deployment infrastructure and tenant company databases:

| Environment | Database Name | Description | Role / Target |
|---|---|---|---|
| **SQL Host** | `100.127.240.115\INSERTGT` | Primary Microsoft SQL Server Instance | Port 1433 |
| **Launcher** | `InsERT_Launcher` | InsERT central catalog storing `.mpkg` binaries | System Catalog |
| **Test Tenant** | `Nexo_Test_1` | Staging database for automated E2E tests and verification | Test / Staging |
| **Demo Tenant** | `Nexo_Demo_1` | Target database for deployment manifest status verification | Verification |
| **Production 1** | `Nexo_KLINEX SP_Z_O_O` | Klinex Sp. z o.o. operational tenant database | Active Production |
| **Production 2** | `Nexo_EUROSUPPORT_GROUP` | Eurosupport Group Sp. z o.o. operational tenant database | Active Production |
| **Production 3** | `Nexo_EUROSUPPORT OUTSOURCING` | Eurosupport Outsourcing Sp. z o.o. operational tenant database | Active Production |

---

## 7. Deployment & Automation Scripts

Production automation scripts reside in [`scripts/`](scripts/):

- **`build_and_package_release.ps1`**: Collects build outputs from `bin\Debug\`, prepares them in project root, and invokes `PackageAndInstallerBuilder.exe` to generate `.mpkg` and standalone `.exe` installers.
- **`install_to_target_dbs.ps1`**: Registers the `.mpkg` package into `InsERT_Launcher`, updates the `DeploymentManifest` on target databases, and triggers binary synchronization. Supports `-QuickSyncOnly`.
- **`sync_dll_to_all_deployments.ps1`**: Hot-copies compiled assemblies directly into all local Gratyfikant client deployment folders (`%LOCALAPPDATA%\InsERT\Deployments\Nexo\`). Supports `-KillRunningNexo`.
- **`check_db_status.ps1`**: Executes a direct SQL query against `InsLauncher.InstalledProducts` on the database server to verify registered package identities.
- **`bump_version.py`**: Synchronizes version strings across `HrAppka Import Pracowników.csproj` and `Konfiguracja.xml`.

*For detailed syntax, parameter descriptions, and console outputs, see [`docs/deployment_and_scripts.md`](docs/deployment_and_scripts.md).*

---

## 8. Maintenance & Git Workflow

- **Repository**: Hosted under version control tracking the `master` and `dev` branches.
- **Stable Releases**: Tagged on `master` following semantic versioning (e.g. `1.0.8.20`).
- **Development**: Feature work occurs on `dev` or dedicated topic branches.
- **Deployment Safety Rule**: Always verify package installation and run simulation tests on `Nexo_Test_1` before executing `install_to_target_dbs.ps1` against production company databases (`Nexo_KLINEX SP_Z_O_O`, `Nexo_EUROSUPPORT_GROUP`, `Nexo_EUROSUPPORT OUTSOURCING`).

---
*Maintained by Antigravity Development Team — November 2026*
