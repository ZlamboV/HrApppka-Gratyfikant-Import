# HrAppka Import Pracowników - Gratyfikant Nexo Pro Plugin

Wtyczka (Plugin) dla systemu kadrowo-płacowego **InsERT Gratyfikant Nexo Pro** umożliwiająca automatyczny import i aktualizację danych pracowników z systemu **HrAppka** na podstawie plików Excel (`.xlsx`).

---

## 1. Struktura i Środowisko Projektu

* **Lokalizacja projektu**: `C:\ NEXO SDK SFERA\HrAppka Import Pracowników`
* **Wersja SDK**: InsERT nexo SDK `60.1.1.9292` (biblioteki binarne w `C:\ NEXO SDK SFERA\nexoSDK_60.1.1.9292\Bin\`)
* **Platforma docelowa**: `.NET 8.0 (Windows 10.0.17763.0)`
* **Technologia**: C# / Sfera zdarzeniowa Gratyfikant Nexo Pro, format zapisu pakietów wdrożeniowych `.mpkg` za pomocą narzędzia `PackageAndInstallerBuilder`.

---

## 2. Pliki i Rola Poszczególnych Klas

| Nazwa pliku | Opis roli i zawartości |
| :--- | :--- |
| **[OperacjaImportuPracownikow.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/OperacjaImportuPracownikow.cs)** | **Główny silnik biznesowy**. Odpowiada za odczyt danych, pre-analizę duplikatów (po PESEL/Paszporcie), weryfikację istniejących umów, zakładanie/aktualizację kartotek pracowników, tworzenie umów zleceń oraz generowanie deklaracji zgłoszeniowych ZUS. |
| **[PracownikImportRow.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/PracownikImportRow.cs)** | Klasa modelu danych wejściowych wczytanych z Excela (mapowanie nazw kolumn, walidacja poprawności formatów przez interfejs `IDataErrorInfo`). |
| **[Program.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/Program.cs)** | Główny punkt wejściowy (`Main`) uruchamiany podczas instalacji/deinstalacji pluginu. Zawiera klasę `AutoInstaller` oraz procedurę lokalnej diagnostyki `SferaRunner`. |
| **[Rozszerzenia.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/Rozszerzenia.cs)** | Metody pomocnicze do dynamicznej ekstrakcji błędów walidacji encji z mechanizmu Sfery (zwracanie szczegółowych komunikatów o niepowodzeniach). |
| **[DepartmentResolver.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/DepartmentResolver.cs)** | Translator i resolver działów/klientów z Excela na struktury organizacyjne (jednostki organizacyjne) w bazie Gratyfikanta. |
| **[PeselHelper.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/PeselHelper.cs)** | Parsowanie PESEL, walidacja sumy kontrolnej, automatyczna ekstrakcja daty urodzenia oraz płci pracownika. |
| **[DostawcaPluginow.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/DostawcaPluginow.cs)** | Metadane producenta rozszerzenia. |
| **[GrupaOperacji.cs](file:///C:/%20NEXO%20SDK%20SFERA/HrAppka%20Import%20Pracownik%C3%B3w/GrupaOperacji.cs)** | Rejestracja wtyczki jako operacji w menu programu Gratyfikant. |

---

## 3. Główne Funkcjonalności Biznesowe

### A. Inteligentna Aktualizacja Kartotek i Unikanie Duplikatów
System dopasowuje wczytanych z Excela pracowników do istniejących w Gratyfikancie po **PESEL** lub numerze **Paszportu**:
* **Nowy pracownik**: Tworzona jest kartoteka osobowa, adres, rachunek bankowy, nowa umowa zlecenie oraz deklaracja ZUS ZUA.
* **Pracownik już istnieje (z taką samą umową)**: 
  * Jeśli w systemie istnieje już `"Umowa zlecenie"` o identycznej dacie rozpoczęcia — wtyczka aktualizuje dane personalne, adres, telefon, NFZ i konto bankowe pracownika, ale **pomija ponowne dodanie umowy** oraz deklarację ZUS (status: `Aktualizacja (Dane)`).
* **Pracownik już istnieje (ale bez umowy na tę datę)**:
  * Aktualizuje dane kartoteki i **dodaje nową umowę zlecenie** wraz z odpowiednią deklaracją ZUS (status: `Aktualizacja (Nowa umowa)`).

### B. Automatyczne Rozwiązywanie Urzędów Skarbowych
Przekazywana w pliku tekstowa nazwa urzędu skarbowego jest oczyszczana z polskich znaków i porównywana z bazą:
1. Wyszukiwanie w lokalnej bazie Nexo.
2. Jeśli brak — wyszukiwanie w bazach szablonów systemowych Nexo po kodzie i nazwie. Brakujący urząd jest **automatycznie rejestrowany** w kartotece.

### C. Mapowanie Państw i Obywatelstw
Obsługiwane są alternatywne zapisy krajów oraz obywatelstw (w tym m.in. poprawne nazewnictwo Płatnika dla obywateli Zimbabwe, Uzbekistanu, Indii czy Wietnamu).

---

## 4. Konfiguracja i Połączenie z Bazami Danych

Plik parametrów instalacji **`ParametryInstalacji.txt`** przechowuje dane połączenia SQL do rejestracji wtyczki:

```ini
Serwer=100.127.240.115\INSERTGT
UwierzytelnianieWindows=Nie
Uzytkownik=sa
Haslo=
BazaDanych=Nexo_Test_GR
```

### Bazy Danych w Systemie:
* **Testowa**: `Nexo_Test_GR`
* **Produkcyjne**:
  - `Nexo_KLINEX SP_Z_O_O`
  - `Nexo_EUROSUPPORT_GROUP`
  - `Nexo_EUROSUPPORT OUTSOURCING`

---

## 5. Kompilacja i Instalacja Wtyczki

Projekt posiada automatyczny skrypt instalacyjny. Aby wdrożyć wtyczkę na wybraną bazę:

1. Otwórz plik **`ParametryInstalacji.txt`** i ustaw docelową bazę w linii `BazaDanych=`.
2. Uruchom skrót **`Instaluj_Automatycznie.bat`** (lub wykonaj poniższe komendy w terminalu):

```powershell
# 1. Kompilacja i budowanie paczki instalatora wtyczki
dotnet build -c Debug

# 2. Rejestracja wtyczki w bazie danych
dotnet run -- --install
```

Po pomyślnym wykonaniu komendy wtyczka zostanie załadowana do serwera launcher i podłączona do wybranej bazy danych. W programie Gratyfikant Nexo pojawi się w menu operacji importu.

---

## 6. Model Pracy z Gitem

Repozytorium wdrożone na GitHub pod adresem: `https://github.com/ZlamboV/HrApppka-Gratyfikant-Import`

* **`master`** — stabilne wersje produkcyjne.
* **`dev`** — aktualna gałąź deweloperska.

Przy wprowadzaniu kolejnych poprawek pamiętaj o przetestowaniu na bazie `Nexo_Test_GR` przed instalacją na bazy produkcyjne.
