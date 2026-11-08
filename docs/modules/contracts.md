# Contract Creation Service & BO DataDomain Deep-Dive

**Source Files**: `Services/ContractCreationService.cs` (contains `IContractCreationService`, `SzablonUmowyItem`, and `ContractCreationService`)  
**Target Framework**: `.NET 8.0` (`net8.0-windows10.0.17763.0`)  
**SDK Dependency**: InsERT Gratyfikant Nexo Pro SDK 62.0.0.9537  
**Namespace**: `HrAppka_Import_Pracowników.Services`, `InsERT.Moria.Kadry.Duze`, `InsERT.Moria.ModelDanych`, `InsERT.Moria.Sfera`  

---

## Table of Contents
1. [Overview & Architectural Purpose](#1-overview--architectural-purpose)
2. [Service Interface & Contracts](#2-service-interface--contracts)
3. [Sfera Contract Lifecycle & Template Management](#3-sfera-contract-lifecycle--template-management)
   - [3.1 Template Enumeration (`ISzablonyUmowPracowniczych`)](#31-template-enumeration-iszablonyumowpracowniczych)
   - [3.2 Dynamic Contract Creation (`IUmowyPracownicze.Utworz`)](#32-dynamic-contract-creation-iumowypracowniczeutworz)
4. [Solving the Multi-Context EF Detached Entity Collision](#4-solving-the-multi-context-ef-detached-entity-collision)
   - [4.1 Root Cause: Independent ObjectContexts per Business Object](#41-root-cause-independent-objectcontexts-per-business-object)
   - [4.2 Dynamic Reflection Strategy for `ModelDanychContainer`](#42-dynamic-reflection-strategy-for-modeldanychcontainer)
   - [4.3 Entity Re-Querying and Attachment Pattern](#43-entity-re-querying-and-attachment-pattern)
5. [Mox Child Entity Sponsorship Rules](#5-mox-child-entity-sponsorship-rules)
   - [5.1 The Unsponsored Modification Trap](#51-the-unsponsored-modification-trap)
   - [5.2 Correct Attachment Order for `StanowiskoWUmowiePracowniczej`](#52-correct-attachment-order-for-stanowiskowumowiepracowniczej)
6. [Civil Contracts vs Employment Contracts](#6-civil-contracts-vs-employment-contracts)
   - [6.1 Type Differentiation & Calendar Handling](#61-type-differentiation--calendar-handling)
   - [6.2 Working Time Schedules & Full-Time Equivalents (FTE)](#62-working-time-schedules--full-time-equivalents-fte)
7. [Gratyfikant Wage Component Alignment Rule](#7-gratyfikant-wage-component-alignment-rule)
   - [7.1 The 1st-of-Month Nexo Constraint](#71-the-1st-of-month-nexo-constraint)
   - [7.2 Iterative Wage Component Population](#72-iterative-wage-component-population)
8. [Student Under 26 Exemption Rules](#8-student-under-26-exemption-rules)
   - [8.1 Exact Chronological Age Evaluation](#81-exact-chronological-age-evaluation)
   - [8.2 ZUS Report & Contribution Suppression](#82-zus-report--contribution-suppression)
   - [8.3 Custom Flag Assignment (`NIEOPODATKOWANE` vs `OPODATKOWANE`)](#83-custom-flag-assignment-nieopodatkowane-vs-opodatkowane)
9. [Comprehensive Reference Catalog & Code Examples](#9-comprehensive-reference-catalog--code-examples)

---

## 1. Overview & Architectural Purpose

`ContractCreationService` encapsulates the business rules and Entity Framework lifecycle requirements for generating worker contracts within **InsERT Gratyfikant Nexo Pro**.

In Gratyfikant Nexo, contracts are complex, multi-entity business objects containing payroll parameters, tax declarations, working-time calendars, organizational position assignments, and ZUS contribution flags. Creating contracts programmatically via Sfera requires strict adherence to:
- Sfera template resolution and dynamic factory creation.
- Multi-context Entity Framework alignment to prevent `IEntityChangeTracker` collisions.
- Mox DataDomain sponsorship rules when creating child entities.
- Nexo-specific wage component validity boundaries.
- Polish tax and social security exemptions for students under 26 years of age.

---

## 2. Service Interface & Contracts

The service is decoupled behind `IContractCreationService` and implemented as a thread-safe singleton (`ContractCreationService.Instance`).

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Namespace: HrAppka_Import_Pracowników.Services
using System.Collections.Generic;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Sfera;

namespace HrAppka_Import_Pracowników.Services
{
    public interface IContractCreationService
    {
        List<SzablonUmowyItem> GetAvailableTemplates(IUchwyt uchwyt);
        
        SzablonUmowyItem? GetDefaultTemplate(IUchwyt uchwyt, string? factoryOrDepartmentName = null);
        
        UmowaPracowniczaGr CreateContract(
            IUchwyt uchwyt,
            Podmiot employeePodmiot,
            PracownikImportRow row,
            int templateId,
            JednostkaOrganizacyjnaGr? department = null,
            decimal? wageRate = null);
    }
}
```

The data transfer object `SzablonUmowyItem` exposes template details to the UI layer:
```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Namespace: HrAppka_Import_Pracowników.Services
// Class: SzablonUmowyItem (ContractCreationService.cs, lines 27-49)
namespace HrAppka_Import_Pracowników.Services
{
    public class SzablonUmowyItem
    {
        public int Id { get; set; }
        public string Nazwa { get; set; } = string.Empty;
        public string Opis { get; set; } = string.Empty;
        public bool IsDefault { get; set; }
        public int TypUmowyId { get; set; }
        public string TypUmowyNazwa { get; set; } = string.Empty;
        public bool IsCivil { get; set; }

        public override string ToString() => Nazwa;
    }
}
```

---

## 3. Sfera Contract Lifecycle & Template Management

### 3.1 Template Enumeration (`ISzablonyUmowPracowniczych`)

Gratyfikant Nexo defines contract templates (`SzablonUmowyPracowniczej`) that preconfigure salary elements, print layouts, and general terms. `ContractCreationService` queries the template manager via Sfera:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Namespace: HrAppka_Import_Pracowników.Services
// Class: ContractCreationService
using System;
using System.Collections.Generic;
using System.Linq;
using InsERT.Moria.Kadry.Duze;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Sfera;

public List<SzablonUmowyItem> GetAvailableTemplates(IUchwyt uchwyt)
{
    var list = new List<SzablonUmowyItem>();
    var szablonyMgr = uchwyt.SzablonyUmowPracowniczychGr();
    if (szablonyMgr == null) return list;

    var query = szablonyMgr.Dane.Wszystkie()
        .Where(s => !s.IsInRecycleBin)
        .OrderBy(s => s.Nazwa);

    foreach (var s in query)
    {
        bool isCivil = s.TypUmowyPracowniczejGr != null &&
                       s.TypUmowyPracowniczejGr.Typ == (byte)TypUmowyPracowniczej.Cywilnoprawna;

        list.Add(new SzablonUmowyItem
        {
            Id = s.Id,
            Nazwa = s.Nazwa ?? $"Szablon {s.Id}",
            Opis = s.Opis ?? string.Empty,
            IsCivilContract = isCivil
        });
    }

    return list;
}
```

---

### 3.2 Dynamic Contract Creation (`IUmowyPracownicze.Utworz`)

Contracts are instantiated dynamically by passing the template entity directly to `IUmowyPracowniczeGr.Utworz`:

```csharp
var szablonyMgr = uchwyt.SzablonyUmowPracowniczychGr();
var templateEntity = szablonyMgr.Dane.Wszystkie().FirstOrDefault(s => s.Id == templateId);

if (templateEntity == null)
{
    throw new InvalidOperationException($"Nie odnaleziono szablonu umowy o ID: {templateId}");
}

var umowyMgr = uchwyt.UmowyPracowniczeGr();
using (IUmowaPracownicza umowaBO = umowyMgr.Utworz(templateEntity))
{
    UmowaPracowniczaGr u = umowaBO.Dane;
    // Configure contract entities...
    
    if (!umowaBO.Zapisz())
    {
        var errors = uchwyt.PodajBledy(umowaBO);
        throw new InvalidOperationException(string.Join("; ", errors.Select(e => e.Tresc)));
    }
    return u;
}
```

---

## 4. Solving the Multi-Context EF Detached Entity Collision

### 4.1 Root Cause: Independent ObjectContexts per Business Object

Every Sfera Business Object (`IObiektBiznesowy`) maintains its own internal Entity Framework `ObjectContext` (`ModelDanychContainer`). If an entity is queried from a different context (e.g. `employeePodmiot` fetched via `uchwyt.Podmioty().Dane.WszystkieOsoby()`) and linked directly to `u.Pracownik`, Entity Framework throws:

```
System.InvalidOperationException:
"An entity object cannot be referenced by multiple instances of IEntityChangeTracker."
```

---

### 4.2 Dynamic Reflection Strategy for `ModelDanychContainer`

To bridge this boundary, `ContractCreationService` extracts the contract BO's own `ModelDanychContainer`:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: ContractCreationService
ModelDanychContainer? context = null;

// Strategy 1: Reflection on 'ObjectContext' property
try
{
    var prop = u.GetType().GetProperty(
        "ObjectContext", 
        System.Reflection.BindingFlags.Instance | 
        System.Reflection.BindingFlags.Public | 
        System.Reflection.BindingFlags.NonPublic);
    context = prop?.GetValue(u) as ModelDanychContainer;
}
catch { }

// Strategy 2: Extract from Mox UnitOfWork Injection Scope
if (context == null)
{
    try
    {
        if (umowaBO is InsERT.Mox.BusinessObjects.IGetUnitOfWork getUow &&
            getUow.UnitOfWork is InsERT.Mox.Runtime.IInjectionScope scope)
        {
            context = scope.ScopedContainer?.GetObject(typeof(ModelDanychContainer)) 
                      as ModelDanychContainer;
        }
    }
    catch { }
}
```

---

### 4.3 Entity Re-Querying and Attachment Pattern

Once the BO's internal `ModelDanychContainer` is obtained, navigation entities are re-queried directly inside that context:

```csharp
// Re-query Employee
var empId = employeePodmiot.Osoba?.Pracownik?.Id;
if (context != null && empId.HasValue && empId.Value > 0)
{
    u.Pracownik = context.Pracownicy.FirstOrDefault(p => p.Id == empId.Value) 
                  ?? employeePodmiot.Osoba?.Pracownik;
}
else
{
    u.Pracownik = employeePodmiot.Osoba?.Pracownik;
}

// Re-query Department
JednostkaOrganizacyjnaGr? attachedDept = null;
if (department != null)
{
    if (context != null && department.Id > 0)
    {
        attachedDept = context.JednostkiOrganizacyjne
            .OfType<JednostkaOrganizacyjnaGr>()
            .FirstOrDefault(j => j.Id == department.Id);
    }
    attachedDept ??= department;
}
```

This ensures all foreign navigation properties belong to the exact same `IEntityChangeTracker`, completely eliminating detachment exceptions.

---

## 5. Mox Child Entity Sponsorship Rules

### 5.1 The Unsponsored Modification Trap

In Gratyfikant Nexo's Mox framework, child collection entities (such as `StanowiskoWUmowiePracowniczej`) are monitored by the parent Business Object's `DataDomain`.

If a developer instantiates a new child entity:
```csharp
var st = new StanowiskoWUmowiePracowniczej();
st.JednostkaOrganizacyjna = attachedDept; // CRASH!
u.Stanowiska.Add(st);
```
Mox throws an `UnsponsoredModificationException` because `st` is modified while not yet sponsored by the parent `DataDomain`.

---

### 5.2 Correct Attachment Order for `StanowiskoWUmowiePracowniczej`

To satisfy Mox sponsorship requirements, the child entity **must be added to the collection first**:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: ContractCreationService
var st = u.Stanowiska.FirstOrDefault();
if (st == null)
{
    st = new StanowiskoWUmowiePracowniczej
    {
        DataOd = dtStart,
        LP = 1
    };
    // CRITICAL: Must be added to parent collection FIRST!
    u.Stanowiska.Add(st);
}

// SAFE: Entity is now sponsored by umowaBO's DataDomain
st.JednostkaOrganizacyjna = attachedDept;
```

---

## 6. Civil Contracts vs Employment Contracts

### 6.1 Type Differentiation & Calendar Handling

Gratyfikant Nexo distinguishes between employment contracts (*Umowa o pracę*) and civil law contracts (*Umowa cywilnoprawna* / *Umowa zlecenie*):

```csharp
bool isCivilContract = templateEntity.TypUmowyPracowniczejGr != null &&
    templateEntity.TypUmowyPracowniczejGr.Typ == (byte)TypUmowyPracowniczej.Cywilnoprawna;
```

- **Working Calendar (`Kalendarz`)**:
  - If the contract type specifies a working time calendar (`u.TypUmowyPracowniczej.Kalendarz == true`), Gratyfikant requires attaching an active calendar schedule. The service searches for `"Bez cyklu"` or uses the template default, configuring `WymiarZatrudnienia = 1/1`.
  - For standard civil contracts (*Umowa zlecenie*), calendars are suppressed: `u.Kalendarze.Clear()`.

---

### 6.2 Working Time Schedules & Full-Time Equivalents (FTE)

When attaching calendars to employment contracts:
```csharp
if (u.TypUmowyPracowniczej.Kalendarz)
{
    var kal = u.Kalendarze.FirstOrDefault();
    if (kal == null)
    {
        kal = new KalendarzWUmowiePracowniczej { DataOd = dtStart, LP = 1 };
        u.Kalendarze.Add(kal);
    }
    
    // Assign "Bez cyklu" calendar
    var defaultKal = context?.Kalendarze.FirstOrDefault(k => k.Nazwa == "Bez cyklu") 
                     ?? uchwyt.Kalendarze()?.Dane?.Wszystkie()?.FirstOrDefault(k => k.Nazwa == "Bez cyklu");
    kal.Kalendarz = defaultKal;
    kal.WymiarLicznik = 1;
    kal.WymiarMianownik = 1;
}
```

---

## 7. Gratyfikant Wage Component Alignment Rule

### 7.1 The 1st-of-Month Nexo Constraint

A strict architectural invariant of Gratyfikant Nexo requires that salary component validity periods (`OkresObowiazywania.DataPoczatkowa`) **must align to the 1st day of the calendar month**:

$$\text{DataPoczatkowa} = \text{new DateTime}(\text{dtStart.Year}, \text{dtStart.Month}, 1)$$

If a contract starts mid-month (e.g. `2026-04-15`), setting `DataPoczatkowa = 2026-04-15` on a `SkladnikPlacowy` will cause Nexo payroll calculations to fail or throw validation exceptions upon contract generation.

---

### 7.2 Iterative Wage Component Population

`ContractCreationService` iterates over all template wage components, updates their rates, and strictly enforces the 1st-of-the-month alignment rule:

```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: ContractCreationService
var componentStart = new DateTime(dtStart.Year, dtStart.Month, 1);

foreach (var sp in u.SkladnikiPlacowe)
{
    foreach (var val in sp.Wartosci)
    {
        val.Wartosc = hasRate ? parsedRate : 0.00m;
        
        if (val.OkresObowiazywania == null)
        {
            val.OkresObowiazywania = new OkresDowolny { DataPoczatkowa = componentStart };
        }
        else
        {
            val.OkresObowiazywania.DataPoczatkowa = componentStart;
        }
    }
}
```

---

## 8. Student Under 26 Exemption Rules

### 8.1 Exact Chronological Age Evaluation

In accordance with Polish tax law (*Ulga dla młodych*) and ZUS regulations, students under 26 years of age working under civil contracts (*umowa zlecenie*) are fully exempt from personal income tax (PIT) and social/health insurance contributions (ZUS).

The service evaluates exact chronological age on the contract start date:

```csharp
DateTime birthDate = row.ResolvedBirthDate ?? DateTime.MinValue;
int age = PeselHelper.CalculateAge(birthDate, dtStart);
bool studentFlag = string.Equals(row.JestStudentem?.Trim(), "Tak", StringComparison.OrdinalIgnoreCase);
bool ageUnder26 = age < 26;
bool isStudentUnder26 = studentFlag && ageUnder26;
bool isZusSubject = !isStudentUnder26;
```

---

### 8.2 ZUS Report & Contribution Suppression

Based on the evaluation:
```csharp
// Assembly: HrAppka_Import_Pracowników.dll
// Class: ContractCreationService
u.RaportZUS = (byte)(isStudentUnder26 ? TypRaportuZUS.Brak : TypRaportuZUS.RCA);

u.NaliczaSkladkiEmerytalne = isZusSubject;
u.NaliczaSkladkiRentowe = isZusSubject;
u.NaliczaSkladkiZdrowotne = isZusSubject;
u.NaliczaSkladkiWypadkowe = isZusSubject;
u.NaliczaSkladkiFP = isZusSubject;
u.NaliczaSkladkiFGSP = isZusSubject;
u.NaliczaSkladkiChorobowe = false; // Voluntary for civil contracts
```

---

### 8.3 Custom Flag Assignment (`NIEOPODATKOWANE` vs `OPODATKOWANE`)

Gratyfikant Nexo tracks taxation categories via custom flags (`FlagaWlasna`):
- Student under 26 -> `"NIEOPODATKOWANE"`
- Normal worker -> `"OPODATKOWANE"`

```csharp
string flagName = isStudentUnder26 ? "NIEOPODATKOWANE" : "OPODATKOWANE";

u.FlagaWlasna = context != null 
    ? context.FlagiWlasne.FirstOrDefault(f => f.Nazwa == flagName)
    : uchwyt.PodajObiektTypu<IFlagiWlasne>()?.Dane?.Wszystkie()?.FirstOrDefault(f => f.Nazwa == flagName);
```

---

## 9. Comprehensive Reference Catalog & Code Examples

### 9.1 API Reference

| Method | Signature | Description |
|---|---|---|
| `GetAvailableTemplates` | `List<SzablonUmowyItem> GetAvailableTemplates(IUchwyt uchwyt)` | Returns active contract templates. |
| `GetDefaultTemplate` | `SzablonUmowyItem? GetDefaultTemplate(IUchwyt uchwyt, string? factoryOrDepartmentName)` | Finds the default civil contract template. |
| `CreateContract` | `UmowaPracowniczaGr CreateContract(IUchwyt, Podmiot, PracownikImportRow, int, JednostkaOrganizacyjnaGr?, decimal?)` | Full contract lifecycle creation and persistence. |

### 9.2 Practical Invocation Example

```csharp
using InsERT.Moria.Sfera;
using HrAppka_Import_Pracowników;
using HrAppka_Import_Pracowników.Services;

public static void GenerateWorkerContract(
    IUchwyt sfera, 
    InsERT.Moria.ModelDanych.Podmiot employeePodmiot, 
    PracownikImportRow importRow)
{
    var contractService = ContractCreationService.Instance;
    
    // 1. Resolve template
    var defaultTemplate = contractService.GetDefaultTemplate(sfera);
    int templateId = importRow.SelectedSzablonUmowy?.Id ?? defaultTemplate!.Id;

    // 2. Resolve department
    var dept = SferaDictionaryCache.Instance.ResolveDepartment(sfera, importRow.DzialNazwa);

    // 3. Create and persist contract
    var contract = contractService.CreateContract(
        uchwyt: sfera,
        employeePodmiot: employeePodmiot,
        row: importRow,
        templateId: templateId,
        department: dept,
        wageRate: 28.10m
    );

    Console.WriteLine($"Utworzono umowę o numerze ID: {contract.Id}");
}
```
