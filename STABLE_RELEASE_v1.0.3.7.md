# 🟢 STABLE RELEASE — HrAppka Import Pracowników v1.0.3.7

> **Статус:** ✅ СТАБІЛЬНА ВЕРСІЯ — НЕ ЗМІНЮВАТИ  
> **Дата релізу:** 2026-07-05  
> **Git Tag:** `v1.0.3.7-stable`  
> **Автор:** Antigravity Dev Team  
> **Copyright:** © 2026 Antigravity Dev Team

---

> [!CAUTION]
> Цей код є стабільною робочою версією. Будь-які зміни повинні вноситися в нову гілку розробки (`dev` або `v2.x`).
> НЕ МОДИФІКУВАТИ файли в гілці `main` / тегу `v1.0.3.7-stable`.

---

## Опис програми

**HrAppka Import Pracowników** — це плагін для **InsERT Gratyfikant Nexo Pro** (платформа Sfera), який автоматизує імпорт працівників з системи HrAppka через Excel-файли (.xlsx).

### Що робить плагін:

1. **Читає Excel-файл** з даними працівників (до 39 колонок)
2. **Валідує дані** — PESEL, паспорти, банківські рахунки, дублікати
3. **Створює працівників** у Gratyfikant Nexo Pro з повними даними:
   - Персональні дані (ім'я, прізвище, PESEL, дата народження, стать)
   - Адреса проживання (повна, з воєводством/повітом/гміною)
   - Контакти (email, телефон)
   - Документи (паспорт із датами)
   - Банківський рахунок (IBAN, валюта, назва банку)
   - Податковий орган (Urząd Skarbowy)
   - Відділ NFZ
4. **Створює договори** (Umowa zlecenie) з повним налаштуванням:
   - ZUS внески (RCA для не-студентів, Brak для студентів < 26)
   - PIT аванси (з урахуванням віку і статусу студента)
   - Графік оплати (щомісячно, зміщення рахунку = 25)
   - Ставки оплати (всі складові = 0.00)
   - Організаційний підрозділ (автостворення під "FABRYKA")
   - Код професії
5. **Підтримує 3 сценарії** створення працівника:
   - Існуюча особа з роллю Pracownik → оновлення
   - Існуюча особа без ролі Pracownik → трансформація
   - Нова особа → створення з нуля
6. **Має авто-інсталятор** — встановлення/видалення через CLI

---

## Архітектура

```
HrAppka_Import_Pracowników (Plugin for Sfera)
│
├── Entry Point
│   ├── DostawcaPluginow.cs    — IDostawcaPluginow (метадані плагіна)
│   └── GrupaOperacji.cs       — IGrupaOperacji (точка входу Sfera)
│
├── Core Logic
│   └── OperacjaImportuPracownikow.cs — Головна операція імпорту (980 рядків)
│       ├── Wykonaj()             — Відкриття файлу, запуск імпорту
│       ├── WczytajExcel()        — Читання Excel через MiniExcel
│       ├── PreAnalyzeRows()      — Пре-валідація (дублікати, PESEL, паспорти)
│       ├── ZaimportujRow()       — Імпорт одного рядка (працівник + договір)
│       └── Helper methods        — Пошук країн, воєводств, NFZ тощо
│
├── Data & Validation
│   ├── PracownikImportRow.cs  — DTO для рядка Excel (31 поле, WPF binding)
│   ├── PeselHelper.cs         — Валідація PESEL, витяг дати/статі
│   └── DepartmentResolver.cs  — Пошук/створення організаційних підрозділів
│
├── Utilities
│   ├── Rozszerzenia.cs        — Extension methods для витягу помилок (UoW)
│   └── BladInfo               — DTO для інформації про помилки
│
└── Installation
    └── Program.cs             — CLI інсталятор + діагностичний раннер
        ├── Program.Main()      — Маршрутизація: --install / --uninstall / діагностика
        ├── AutoInstaller       — Встановлення .mpkg через InsERT Mox Launcher
        └── SferaRunner         — Тестовий харнес (НЕ продакшн)
```

---

## Технічний стек

| Компонент | Значення |
|---|---|
| **Target Framework** | `net8.0-windows10.0.17763.0` |
| **Output Type** | `Exe` (працює як плагін DLL і як CLI інсталятор) |
| **Namespace** | `HrAppka_Import_Pracowników` |
| **WPF** | Enabled (для діалогів і UI операцій) |
| **NuGet залежності** | `MiniExcel` v1.34.0 |
| **Nexo SDK** | v60.1.1.9292 |
| **Build PostTarget** | PackageAndInstallerBuilder.exe → .mpkg + .exe |

### SDK Assembly References

| Assembly | Призначення |
|---|---|
| `InsERT.Mox.Launcher` | Встановлення/видалення пакетів |
| `InsERT.Moria.API` | Основний API Sfera |
| `InsERT.Moria.Sfera` | Підключення до бази Nexo |
| `InsERT.Moria.ModelDanych` | Моделі даних Gratyfikant |
| `InsERT.Moria.DaneDomyslne` | Дефолтні дані |
| `InsERT.Mox.Core` | Ядро Mox платформи |
| `InsERT.Mox.EntityFramework.Core` | EF Core інтеграція |
| `InsERT.Mox.EntityFrameworkSupport` | EF підтримка |
| `Microsoft.Data.SqlClient` | SQL Server з'єднання |

---

## Файли релізу

### Вихідний код (8 файлів)

| Файл | Рядків | Опис |
|---|---|---|
| `OperacjaImportuPracownikow.cs` | 980 | Головна логіка імпорту |
| `Program.cs` | 579 | CLI інсталятор + тестовий раннер |
| `PracownikImportRow.cs` | 387 | DTO для Excel рядка |
| `Rozszerzenia.cs` | 95 | Extension methods для помилок |
| `PeselHelper.cs` | 93 | Валідація PESEL |
| `DepartmentResolver.cs` | 68 | Пошук/створення підрозділів |
| `GrupaOperacji.cs` | 24 | Точка входу плагіна |
| `DostawcaPluginow.cs` | 23 | Метадані постачальника |

### Конфігурація

| Файл | Опис |
|---|---|
| `Konfiguracja.xml` | Конфігурація пакету (версія 1.0.3.7, вміст, інсталятор) |
| `HrAppka_Import_Pracowników.plugin` | Точка входу Sfera → `GrupaOperacji` |
| `ParametryInstalacji.txt` | Параметри підключення до SQL Server |
| `Instaluj_Automatycznie.bat` | Автоматична збірка + інсталяція |

### Артефакти збірки

| Файл | Розмір | Опис |
|---|---|---|
| `bin/Debug/HrAppka_Import_Pracowników.dll` | 70 KB | Плагін DLL |
| `bin/Debug/HrAppka_Import_Pracowników.exe` | 152 KB | CLI інсталятор |
| `bin/Debug/Pakiet/*.mpkg` | 129 KB | Пакет для Mox Launcher |
| `bin/Debug/Instalator/*.exe` | 25 MB | Повний інсталятор |

---

## Відомі обмеження та TODO (для наступної версії)

### Вимкнена функціональність

1. **Громадянство (Obywatelstwo)** — блок коду закоментований (`/* ... */`) в `OperacjaImportuPracownikow.cs`
2. **Декларація ZUS** — блок існує але вимкнений (`if (false && ...)`, рядки 807-880)

### Невідповідності

3. **`Instaluj_Automatycznie.bat`** — echo повідомлення каже "version 1.0.0.2", хоча фактична версія = 1.0.3.7
4. **Графік оплати** — `DataRachunkuPrzesuniecie = 25` і `PrzesuniecieNaliczenia = 0` можливо переплутані відносно специфікації в `project_context.md`
5. **OutputType = Exe** — для чистого продакшн плагіна має бути `Library`, але `Exe` потрібен для CLI інсталятора

### Технічний борг

6. **Хардкоджені облікові дані** в `SferaRunner` (тестовий код) — сервер, пароль SA, оператор
7. **Немає юніт-тестів**
8. **`msbuild.log`** (1.3 MB) лежить у проекті
9. **Хардкоджений шлях до Excel** в `SferaRunner`

---

## Підключення до бази (тестове середовище)

```
Сервер:   100.127.240.115\INSERTGT
Auth:     SQL Server (sa / пустий пароль)
База:     Nexo_Test_GR
Оператор: Szef / robocze
```

---

## GUID плагіна

```
8e2be35b-12d8-4f81-9b16-562a93bbef88
```

---

## Як зібрати і встановити

```bash
# Збірка
dotnet build -c Debug

# Встановлення плагіна в Nexo
dotnet run -- --install

# Видалення плагіна з Nexo
dotnet run -- --uninstall
```

Або через `Instaluj_Automatycznie.bat` для повної автоматизації.

---

## Правила роботи з цим релізом

1. ❌ **НЕ змінювати** файли в тегу `v1.0.3.7-stable`
2. ✅ **Створити нову гілку** (`dev`, `v2.0`, etc.) для подальшої розробки
3. ✅ **Використовувати цю версію** як reference при розробці нових фіч
4. ✅ **Цей інсталятор** (`HrAppka_Import_Pracowników-1.0.3.7.exe`) можна деплоїти на продакшн
