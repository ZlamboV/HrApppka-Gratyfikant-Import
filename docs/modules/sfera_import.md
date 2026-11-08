# Sfera Native 2.0 Import Engine Deep-Dive

**Source Files**: `OperacjaImportuPracownikow20.cs`, `GrupaOperacji.cs`, `DostawcaPluginow.cs`, `Rozszerzenia.cs`, `Services/SferaDictionaryCache.cs`  
**Target Framework**: `.NET 8.0` (`net8.0-windows10.0.17763.0`)  
**SDK Dependency**: InsERT Gratyfikant Nexo Pro SDK 62.0.0.9537  
**Namespace**: `HrAppka_Import_Pracowników`, `HrAppka_Import_Pracowników.Services`  

---

## Table of Contents
1. [Overview & Architectural Role](#1-overview--architectural-role)
2. [Plugin Registration & Sfera Contracts](#2-plugin-registration--sfera-contracts)
   - [2.1 Plugin Provider (`DostawcaPluginow.cs`)](#21-plugin-provider-dostawcapluginowcs)
   - [2.2 Operation Group (`GrupaOperacji.cs`)](#22-operation-group-grupaoperacjics)
   - [2.3 Operation Binding (`OperacjaNaLiscieDanych`)](#23-operation-binding-operacjonalisciedanych)
3. [Employee Creation & Update Lifecycle](#3-employee-creation--update-lifecycle)
   - [3.1 Branching Strategy: New vs Existing Employee Cards](#31-branching-strategy-new-vs-existing-employee-cards)
   - [3.2 Address Assignment & Spatial Hierarchies](#32-address-assignment--spatial-hierarchies)
   - [3.3 Bank Account Collision Recovery Pattern](#33-bank-account-collision-recovery-pattern)
   - [3.4 NFZ Branch Mapping & Deployment Month Rules](#34-nfz-branch-mapping--deployment-month-rules)
   - [3.5 Personal Identity & Foreign Name Normalization](#35-personal-identity--foreign-name-normalization)
4. [Duplicate Suppression & Conflict Resolution Gate](#4-duplicate-suppression--conflict-resolution-gate)
   - [4.1 Identification of Existing Entities](#41-identification-of-existing-entities)
   - [4.2 The Four Conflict Resolution Actions](#42-the-four-conflict-resolution-actions)
   - [4.3 Active Contract Termination on Overwrite](#43-active-contract-termination-on-overwrite)
   - [4.4 Exact Contract Duplicate Prevention](#44-exact-contract-duplicate-prevention)
5. [Deep Error Diagnostics via Rozszerzenia.cs](#5-deep-error-diagnostics-via-rozszerzeniacs)
   - [5.1 The Limitations of Standard Sfera Errors](#51-the-limitations-of-standard-sfera-errors)
   - [5.2 Mox DataDomain & Unit of Work Traversal Implementation](#52-mox-datadomain--unit-of-work-traversal-implementation)
6. [High-Performance Pre-Caching via SferaDictionaryCache](#6-high-performance-pre-caching-via-sferadictionarycache)
   - [6.1 Cache Initialization & Lifecycle](#61-cache-initialization--lifecycle)
   - [6.2 Anonymous Database Projections for Existing People](#62-anonymous-database-projections-for-existing-people)
   - [6.3 Dynamic Tax Office Auto-Creation from Master Templates](#63-dynamic-tax-office-auto-creation-from-master-templates)
   - [6.4 Hierarchical Department Provisioning](#64-hierarchical-department-provisioning)
7. [API & Method Reference Catalog](#7-api--method-reference-catalog)

---

## 1. Overview & Architectural Role

`OperacjaImportuPracownikow20` represents the primary orchestration engine for the Sfera Native 2.0 import architecture. Designed to replace the legacy monolithic v1 import (`OperacjaImportuPracownikow.cs`), v2 decouples data ingestion, in-memory caching, interactive validation, contract creation, and ZUS declaration persistence.

The engine hooks directly into the **Gratyfikant Nexo Pro** user interface as a native plugin operation, executing in-process with access to Nexo's `IUchwyt` handle and Microsoft SQL Server database.

---

## 2. Plugin Registration & Sfera Contracts

### 2.1 Plugin Provider (`DostawcaPluginow.cs`)

InsERT Nexo discovers plugins by querying types implementing `InsERT.Moria.Rozszerzanie.IDostawcaPluginow`. This class exposes vendor metadata and contact details:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Namespace: HrAppka_Import_Pracowników
using System.Collections.Generic;
using InsERT.Moria.Rozszerzanie;

namespace HrAppka_Import_Pracowników
{
    public class DostawcaPluginow : IDostawcaPluginow
    {
        public string Nazwa => "Antigravity Dev Team";
        public string Adres => "Kijów, Ukraina";
        public string AdresWWW => "www.example.com";
        public IEnumerable<string> Kontakty => new string[] { "support@example.com" };
        public string KRS => "";
        public string NIP => "";
        public string REGON => "";
    }
}
```

### 2.2 Operation Group (`GrupaOperacji.cs`)

Operation groups partition plugin actions within Nexo's menu hierarchy. `GrupaOperacji` implements `InsERT.Moria.Rozszerzanie.IGrupaOperacji` with a unique GUID:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Namespace: HrAppka_Import_Pracowników
using System;
using System.Collections.Generic;
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Rozszerzanie.Operacje;

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
            new OperacjaImportuPracownikow20()
        };
    }
}
```

> **Design Decision**: Notice that `GrupaOperacji.Operacje` registers **only** `OperacjaImportuPracownikow20`. Legacy `OperacjaImportuPracownikow` (v1) is omitted, isolating users from obsolete code.

### 2.3 Operation Binding (`OperacjaNaLiscieDanych`)

`OperacjaImportuPracownikow20` inherits from `InsERT.Moria.Rozszerzanie.Operacje.OperacjaNaLiscieDanych<Pracownik, int>`:
- **Target List**: Binds to the Gratyfikant Nexo "Pracownicy" (Employees) module view.
- **Menu Placement**: `SciezkaWMenu => null` places the button directly in the standard context action ribbon and the "Operacje" dropdown.
- **Display Caption**: `Nazwa => "Import Pracowników"`.
- **Pre-execution Gate**: `SprawdzCzyMoznaWykonac(...) => true` ensures the button remains clickable regardless of whether the user has selected 0, 1, or multiple rows in the Nexo grid.

---

## 3. Employee Creation & Update Lifecycle

### 3.1 Branching Strategy: New vs Existing Employee Cards

The method `ZaimportujRow20` determines the appropriate business object lifecycle based on duplicate analysis and the operator's decision (`DecyzjaIstniejacyPracownik`):

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: OperacjaImportuPracownikow20
// Method: ZaimportujRow20

IPodmiot? pracownikBO = null;
Podmiot? savedEmployeePodmiot = null;

if (existingOsoba != null && item.DecyzjaDlaIstniejacych == DecyzjaIstniejacyPracownik.TylkoDodajUmowe)
{
    // Path A: Preserve existing employee card
    if (existingOsoba.Pracownik == null)
    {
        // If entity exists as a client/person but lacks employee role, transform it
        savedEmployeePodmiot = podmiotyMgr.PrzeksztalcOsobeNaPracownika(existingOsoba);
    }
    else
    {
        savedEmployeePodmiot = existingOsoba.Podmiot;
    }
}
else if (existingOsoba != null)
{
    // Path B: Update existing employee card
    if (existingOsoba.Pracownik != null)
    {
        pracownikBO = podmiotyMgr.Znajdz(existingOsoba);
    }
    else
    {
        var transformed = podmiotyMgr.PrzeksztalcOsobeNaPracownika(existingOsoba);
        pracownikBO = podmiotyMgr.Znajdz(transformed);
    }
}
else
{
    // Path C: Brand new employee card
    pracownikBO = podmiotyMgr.UtworzPracownika();
    pracownikBO.AutoSymbol();
}
```

---

### 3.2 Address Assignment & Spatial Hierarchies

Address persistence in Gratyfikant Nexo requires synchronizing multiple fields across `Podmiot.AdresPodstawowy`:

1. **Detailed Structure (`Szczegoly`)**: Populates `Ulica`, `NrDomu`, `NrLokalu`, `Miejscowosc`, `KodPocztowy`, and `Poczta`.
2. **Country Association**: Sets `AdresPodstawowy.Panstwo` using `SferaDictionaryCache.Instance.CountriesByCode["PL"]` or foreign country records.
3. **Compound Text Lines**: Formulates compound display strings:
   - `Linia1`: `"ul. {Ulica} {NrDomu}/{NrLokalu}"`
   - `Linia2`: `"{KodPocztowy} {Miejscowosc}"`
   - `LiniaCalosc`: Formatted full mailing address.
4. **Administrative Hierarchy**: Links `Wojewodztwo`, `Powiat`, and `Gmina` using values resolved by `AddressEnrichmentService`.

---

### 3.3 Bank Account Collision Recovery Pattern

In recruitment databases, workers frequently provide invalid or shared bank account numbers (e.g. an account already registered to another contractor, spouse, or vendor in Nexo). 

When Sfera encounters an existing IBAN assigned to a different subject, `pracownikBO.Zapisz()` fails with a validation error containing `"rachunek bankowy"`.

`OperacjaImportuPracownikow20` implements automated recovery:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: OperacjaImportuPracownikow20
bool saveSuccess = pracownikBO.Zapisz();

if (!saveSuccess)
{
    var errs = uchwyt.PodajBledy(pracownikBO);
    string errText = string.Join("; ", errs.Select(e => e.Tresc));

    // Collision Detection: Check if failure was caused by duplicate IBAN
    if (errText.IndexOf("rachunek bankowy", StringComparison.OrdinalIgnoreCase) >= 0 ||
        errText.IndexOf("rachunku bankowego", StringComparison.OrdinalIgnoreCase) >= 0)
    {
        LogNexo.Ostrzezenie($"Wykryto kolizję rachunku bankowego dla pracownika {item.PracownikCol}. Czyszczenie rachunku i ponawianie zapisu.");
        
        // Wipe bank accounts on the employee card to clear the collision
        pracownikBO.Dane.Rachunki.Clear();
        pracownikBO.Dane.RachunekPodstawowy = null;
        
        item.Ostrzezenia += "Rachunek bankowy koliduje z innym podmiotem w bazie – kartotekę zapisano bez rachunku bankowego. ";

        // Retry saving without the problematic bank account
        saveSuccess = pracownikBO.Zapisz();
    }
}
```

This prevents the failure of an entire employee record import due to a non-blocking banking conflict.

---

### 3.4 NFZ Branch Mapping & Deployment Month Rules

Gratyfikant Nexo requires specific social and health insurance initial parameters:

1. **NFZ Branch**: Mapped via `MappedNFZ(item.NfzStr)` from string inputs (e.g. `"01"`, `"02"`, or `"Dolnośląski"`) to the internal Gratyfikant enumeration `OddzialNFZ`.
2. **Accession Date**: `osoba.Pracownik.DataPrzystapieniaDoNFZ = contractStart`.
3. **Deployment Month Rule (`MiesiacPierwszejWyplaty`)**:
   Gratyfikant Nexo tracks the initial payroll generation month via `DaneWdrozenioweGr`:
   ```csharp
   pracownikGr.DaneWdrozenioweGr.MiesiacPierwszejWyplaty = 
       new DateTime(contractStart.Year, contractStart.Month, 1);
   ```

---

### 3.5 Personal Identity & Foreign Name Normalization

1. **Foreign Diacritics**: Replaces international characters commonly found in foreign worker records:
   - `ñ` / `Ñ` -> `n` / `N`
   - `ü` / `Ü` -> `u` / `U`
   - `ä` / `Ä` -> `a` / `A`
   - `ß` -> `ss`
2. **Second Name Discrepancy Gate**:
   If column 0 contains a second name (e.g. `"JAN PAWEŁ"`) while column 7 explicitly specifies another (e.g. `"PIOTR"`), the system checks if UI interaction is permitted. If so, it invokes `ZapytajOUzytkownikaODrugieImie` or concatenates them (`"PAWEŁ PIOTR"`).

---

## 4. Duplicate Suppression & Conflict Resolution Gate

### 4.1 Identification of Existing Entities

During in-memory pre-analysis (`PreAnalyzeRows`), incoming rows are matched against `SferaDictionaryCache`:
1. **PESEL Match**: Checks `cache.ExistingByPesel.TryGetValue(row.Pesel.Trim(), out var peselMatch)`.
2. **Passport Match**: If PESEL is missing or unverified, checks `cache.ExistingByPassport.TryGetValue(row.Paszport.Trim().ToUpperInvariant(), out var passMatch)`.

If either match is found, the row is marked `IsDuplicate = true`, `Status = "Istnieje w bazie"`, and assigned default resolution `AktualizujKartotekeIDodajUmowe`.

---

### 4.2 The Four Conflict Resolution Actions

When an employee already exists in Nexo, `OperacjaImportuPracownikow20` routes execution according to `item.DecyzjaDlaIstniejacego`:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Enum: DecyzjaIstniejacyPracownik
public enum DecyzjaIstniejacyPracownik
{
    AktualizujKartotekeIDodajUmowe, // Updates card data & creates parallel contract
    TylkoDodajUmowe,                // Keeps existing card untouched & creates contract
    NadpiszUmowe,                   // Updates card, closes prior active contracts, creates contract
    PominWiersz                     // Completely ignores row
}
```

---

### 4.3 Active Contract Termination on Overwrite

When the operator selects `NadpiszUmowe`, any existing active contracts for the employee must be cleanly closed before instantiating the new contract:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: OperacjaImportuPracownikow20
if (item.DecyzjaDlaIstniejacego == DecyzjaIstniejacyPracownik.NadpiszUmowe &&
    savedEmployeePodmiot.Osoba?.Pracownik != null)
{
    var existingContracts = uchwyt.UmowyPracowniczeGr().Dane.Wszystkie()
        .Where(x => x.Pracownik != null && x.Pracownik.Id == savedEmployeePodmiot.Osoba.Pracownik.Id)
        .ToList();

    foreach (var existingContract in existingContracts)
    {
        // Check if contract is active (open-ended or ends after new contract start date)
        if (existingContract.OkresObowiazywania != null &&
            (!existingContract.OkresObowiazywania.DataKoncowa.HasValue ||
             existingContract.OkresObowiazywania.DataKoncowa.Value >= dtStart))
        {
            using (var umowaBO = uchwyt.UmowyPracowniczeGr().Znajdz(existingContract))
            {
                // Terminate contract on the day preceding the new start date
                umowaBO.Dane.OkresObowiazywania.DataKoncowa = dtStart.AddDays(-1);
                if (umowaBO.Zapisz())
                {
                    LogNexo.Informacja($"Zamknięto aktywną umowę ID={existingContract.Id} z dniem {dtStart.AddDays(-1):yyyy-MM-dd}.");
                }
            }
        }
    }
}
```

---

### 4.4 Exact Contract Duplicate Prevention

If an import is run repeatedly with the same file, creating identical contracts would pollute employee records. `OperacjaImportuPracownikow20` verifies that an exact contract (same employee, same start date, same end date) does not already exist:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: OperacjaImportuPracownikow20
DateTime? dtEndVal = DateTime.TryParse(item.DataZakonczenia, out var parsedDe) ? parsedDe : null;

var exactMatch = uchwyt.UmowyPracowniczeGr().Dane.Wszystkie()
    .Where(x => x.Pracownik != null && x.Pracownik.Id == savedEmployeePodmiot.Osoba.Pracownik.Id)
    .AsEnumerable()
    .Where(x => !x.IsInRecycleBin)
    .FirstOrDefault(x => x.OkresObowiazywania != null &&
                         x.OkresObowiazywania.DataPoczatkowa == dtStart &&
                         x.OkresObowiazywania.DataKoncowa == dtEndVal);

if (exactMatch != null)
{
    item.Ostrzezenia += $"Umowa na ten okres już istnieje w bazie (ID={exactMatch.Id}). ";
    return WynikOperacji.ZakonczonaPowodzeniem(
        $"Pominięto tworzenie umowy - umowa na ten okres już istnieje (ID={exactMatch.Id}).");
}
```

---

## 5. Deep Error Diagnostics via Rozszerzenia.cs

### 5.1 The Limitations of Standard Sfera Errors

When `IObiektBiznesowy.Zapisz()` returns `false`, calling standard exception handlers often yields vague messages like:
> `"Zapis obiektu nie powiódł się. Sprawdź błędy walidacji."`

Gratyfikant Nexo utilizes an internal validation store (`IValidationMetadataStore`) and Unit of Work participant graph. Member-level validation failures on nested child entities (such as an invalid bank branch, a missing tax office field, or an invalid postal code) are stored in individual participant change trackers.

---

### 5.2 Mox DataDomain & Unit of Work Traversal Implementation

The extension method `Rozszerzenia.PodajBledy` traverses the entire Unit of Work graph and Mox `DataDomain` to aggregate member-level error details:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Namespace: HrAppka_Import_Pracowników
// Class: Rozszerzenia & BladInfo
using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using InsERT.Moria.Sfera;
using InsERT.Moria.Rozszerzanie;
using InsERT.Mox.Validation;
using InsERT.Mox.ObiektyBiznesowe;
using InsERT.Mox.BusinessObjects;
using InsERT.Mox.DataAccess;

namespace HrAppka_Import_Pracowników
{
    // Klasa pomocnicza przechowująca szczegółowe informacje o błędzie walidacji Mox
    internal class BladInfo
    {
        public string Tresc { get; private set; }
        public DataErrorSeverity Waznosc { get; private set; }

        public BladInfo(string tresc, DataErrorSeverity waznosc)
        {
            Tresc = tresc;
            Waznosc = waznosc;
        }
    }

    // Metody rozszerzające IUchwyt oraz IBusinessObject dla diagnostyki błędów Sfery
    internal static class Rozszerzenia
    {
        // Główna metoda rozszerzająca Uchwyt Sfery
        internal static BladInfo[] PodajBledy(this IUchwyt sfera, IObiektBiznesowy obiektBiznesowy)
        {
            var lista = new List<BladInfo>();
            var store = sfera.PodajObiektTypu<IValidationMetadataStore>();

            // 1. Walidacja głównego obiektu biznesowego
            PodajBledy((IBusinessObject)obiektBiznesowy, store, lista);

            // 2. Walidacja wszystkich uczestników Unit of Work
            var uow = ((IGetUnitOfWork)obiektBiznesowy).UnitOfWork;
            foreach (var innyObiektBiznesowy in uow.Participants.OfType<IBusinessObject>().Where(bo => bo != obiektBiznesowy))
            {
                PodajBledy(innyObiektBiznesowy, store, lista);
            }

            return lista.ToArray();
        }

        // Metoda pomocnicza sprawdzająca występowanie błędów na encji
        private static bool HasAnyErrors(IValidationMetadataStore store, ITypedDataErrorInfo errorInfo)
        {
            return errorInfo != null && (errorInfo.Errors.Any() || errorInfo.MemberErrors.Any());
        }

        // Metoda pomocnicza zbierająca szczegółowe informacje o błędach z DataDomain
        private static void PodajBledy(this IBusinessObject obiektBiznesowy,
            IValidationMetadataStore store,
            List<BladInfo> bledy)
        {
            HashSet<ITypedDataErrorInfo> invalidData = new HashSet<ITypedDataErrorInfo>();

            ((IGetDataDomain)obiektBiznesowy).DataDomain.TraverseData(
                obiektBiznesowy.Data, (o) =>
                {
                    ITypedDataErrorInfo? dataErrorInfoEx = o as ITypedDataErrorInfo;
                    if (dataErrorInfoEx != null && HasAnyErrors(store, dataErrorInfoEx))
                    {
                        invalidData.Add(dataErrorInfoEx);
                    }
                    return true;
                });

            foreach (var encjaZBledami in invalidData)
            {
                // Błędy na poziomie całej encji
                foreach (var bladNaCalejEncji in encjaZBledami.Errors)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(bladNaCalejEncji.ToString());
                    sb.Append(" na encjach:" + encjaZBledami.GetType().Name);
                    DataErrorType errorType = store.GetEntryForClrType(bladNaCalejEncji.GetType());
                    bledy.Add(new BladInfo(sb.ToString(), errorType.Severity));
                }

                // Błędy na poziomie konkretnych pól encji
                foreach (var bladNaKonkretnychPolach in encjaZBledami.MemberErrors)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(bladNaKonkretnychPolach.Key.ToString());
                    sb.AppendLine(" na polach:");
                    sb.AppendLine(string.Join(", ", bladNaKonkretnychPolach.Select(b => encjaZBledami.GetType().Name + "." + b)));
                    DataErrorType errorType = store.GetEntryForClrType(bladNaKonkretnychPolach.Key.GetType());
                    bledy.Add(new BladInfo(sb.ToString(), errorType.Severity));
                }
            }
        }
    }
}
```

This guarantees full visibility into validation rejections during automated batch runs.

---

## 6. High-Performance Pre-Caching via SferaDictionaryCache

### 6.1 Cache Initialization & Lifecycle

`SferaDictionaryCache` is implemented as a thread-safe singleton (`SferaDictionaryCache.Instance`). It is loaded during Step 2 of `OperacjaImportuPracownikow20.Wykonaj()`:

```csharp
SferaDictionaryCache.Instance.Initialize(uchwyt);
```

Dictionaries maintained in-memory:
- `CountriesByCode` (`Dictionary<string, Panstwo>`)
- `CountriesByName` (`Dictionary<string, Panstwo>`)
- `CitizenshipsByName` (`Dictionary<string, Obywatelstwo>`)
- `TaxOfficesByCode` (`Dictionary<string, OrganPodatkowy>`)
- `TaxOfficesByName` (`Dictionary<string, OrganPodatkowy>`)
- `AllTaxOffices` (`List<OrganPodatkowy>`)
- `AllMasterUrzedy` (`List<UrzadSkarbowy>`)
- `DepartmentsByName` (`Dictionary<string, JednostkaOrganizacyjnaGr>`)
- `ExistingByPesel` (`Dictionary<string, ExistingPersonInfo>`)
- `ExistingByPassport` (`Dictionary<string, ExistingPersonInfo>`)
- `CurrenciesBySymbol` (`Dictionary<string, Waluta>`)
- `FlagsByName` (`Dictionary<string, FlagaWlasna>`)

---

### 6.2 Anonymous Database Projections for Existing People

Fetching all existing employees via `WszystkieOsoby().ToList()` causes massive memory overhead. `SferaDictionaryCache` uses a lightweight projection:

```csharp
var existingRaw = uchwyt.Podmioty().Dane.WszystkieOsoby()
    .Where(p => p.Osoba != null)
    .Select(p => new
    {
        PodmiotId = p.Id,
        p.Osoba.PESEL,
        p.Osoba.SeriaINumerDokumentuTozsamosci,
        p.Osoba.TypDokumentuTozsamosci,
        IsPracownik = p.Osoba.Pracownik != null
    })
    .ToList();
```

---

### 6.3 Dynamic Tax Office Auto-Creation from Master Templates

If a worker's spreadsheet specifies an official Polish tax office (e.g. `"Urząd Skarbowy w Końskich"`) that exists in Nexo's master template directory (`IUrzedySkarboweDane`) but has never been registered in the client's company database, `ResolveTaxOffice` creates it dynamically:

```csharp
using (var organBO = podmiotyMgr.UtworzOrganPodatkowy())
{
    if (organBO is IOrganPodatkowy organInterface && 
        organInterface.WypelnijDlaKodu(exactTemplate.Kod))
    {
        organBO.AutoSymbol();
        if (organBO.Zapisz())
        {
            var createdOrgan = organBO.Dane.Firma.OrganPodatkowy;
            AllTaxOffices.Add(createdOrgan);
            TaxOfficesByCode[exactTemplate.Kod] = createdOrgan;
            TaxOfficesByName[key] = createdOrgan;
            return createdOrgan;
        }
    }
}
```

---

### 6.4 Hierarchical Department Provisioning

To prevent contract import failures caused by missing organizational units, `SferaDictionaryCache.ResolveDepartment` ensures that:
1. Root department `"FABRYKA"` is verified/created.
2. Target subdepartment (e.g. `"MONTAŻ"`) is verified or dynamically created with `JednostkaNadrzedna = RootFabrykaDepartment`.

---

## 7. API & Method Reference Catalog

| Class / Method | Signature | Description |
|---|---|---|
| `OperacjaImportuPracownikow20.Wykonaj` | `protected override void Wykonaj(IReadOnlyCollection<int> identyfikatoryWybranychElementow, IKontekstListyDanych kontekstListyDanych, IKontekstOperacji kontekstOperacji)` | Sfera SDK entrypoint; launches Excel dialog, warmup, WPF preview, and worker execution. |
| `OperacjaImportuPracownikow20.ZaimportujRow20` | `public static WynikOperacji ZaimportujRow20(IKontekstOperacji? kontekstOperacji, PracownikImportRow item, IUchwyt? fallbackUchwyt = null, List<ZusRegistrationRequest>? pendingZusBatch = null)` | Processes a single employee row: creates/updates card, sets tax office, creates contract, and queues ZUS requests. |
| `OperacjaImportuPracownikow20.WczytajExcel` | `private static List<PracownikImportRow> WczytajExcel(string sciezka)` | Delegates file ingestion to `UniversalFileLoader.LoadFile()`. |
| `SferaDictionaryCache.Initialize` | `public void Initialize(IUchwyt uchwyt)` | Pre-caches reference tables and projects existing persons to eliminate N+1 queries. |
| `SferaDictionaryCache.ResolveTaxOffice` | `public OrganPodatkowy? ResolveTaxOffice(IUchwyt uchwyt, string? usName)` | 4-tier tax office resolution with template auto-creation. |
| `SferaDictionaryCache.ResolveDepartment` | `public JednostkaOrganizacyjnaGr? ResolveDepartment(IUchwyt uchwyt, string? deptName)` | Provisions and links departments under root `"FABRYKA"`. |
| `Rozszerzenia.PodajBledy` | `internal static BladInfo[] PodajBledy(this IUchwyt sfera, IObiektBiznesowy bo)` | Traverses Mox DataDomain and active Unit of Work to extract member validation errors. |
