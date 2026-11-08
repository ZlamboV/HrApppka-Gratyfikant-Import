using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using MiniExcelLibs;

namespace HrAppka_Import_Pracowników.Services
{
    public interface IUniversalFileLoader
    {
        List<PracownikImportRow> Load(string filePath);
    }

    /// <summary>
    /// Loader plików importu pracowników obsługujący pliki .xlsx (poprzez MiniExcel) oraz .csv (poprzez Rfc4180CsvReader).
    /// Posiada stały, zablokowany układ kolumn (Fixed Standard Layout) oparty na wzorcu nowej platformy (36 kolumn),
    /// co całkowicie eliminuje pomyłki i przesunięcia kolumn mogące wynikać z podobnych słów w nagłówkach.
    /// </summary>
    public class UniversalFileLoader : IUniversalFileLoader
    {
        private readonly IRfc4180CsvReader _csvReader;

        public UniversalFileLoader(IRfc4180CsvReader? csvReader = null)
        {
            _csvReader = csvReader ?? new Rfc4180CsvReader();
        }

        public static List<PracownikImportRow> LoadFile(string filePath)
        {
            return new UniversalFileLoader().Load(filePath);
        }

        public List<PracownikImportRow> Load(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Ścieżka do pliku nie może być pusta.", nameof(filePath));
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Wskazany plik nie istnieje: {filePath}", filePath);

            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            List<IReadOnlyList<string>> rawRows;

            if (ext == ".csv" || ext == ".txt")
            {
                rawRows = _csvReader.StreamRows(filePath).ToList();
            }
            else if (ext == ".xlsx" || ext == ".xls")
            {
                rawRows = ReadExcelRows(filePath);
            }
            else
            {
                throw new NotSupportedException($"Format pliku '{ext}' nie jest obsługiwany. Dozwolone rozszerzenia: .xlsx, .csv.");
            }

            if (rawRows.Count < 2)
            {
                return new List<PracownikImportRow>();
            }

            var headerRow = rawRows[0];
            if (headerRow.Count < 15)
            {
                throw new InvalidOperationException($"Plik ma nieprawidłową liczbę kolumn ({headerRow.Count}). Oczekiwano pliku z danymi pracowników.");
            }

            // Rozpoznanie czy to stary format 38-kolumnowy (archiwalna HrAppka z kolumną 'Typ umowy' na poz. 5)
            bool isLegacy38 = headerRow.Count >= 38 &&
                              (string.Equals(headerRow[1]?.Trim(), "Dział", StringComparison.OrdinalIgnoreCase) ||
                               string.Equals(headerRow[4]?.Trim(), "Typ umowy", StringComparison.OrdinalIgnoreCase));

            var rows = new List<PracownikImportRow>();

            for (int r = 1; r < rawRows.Count; r++)
            {
                var values = rawRows[r];

                string GetCol(int zeroBasedIdx)
                {
                    if (zeroBasedIdx >= 0 && zeroBasedIdx < values.Count)
                    {
                        return values[zeroBasedIdx]?.Trim() ?? "";
                    }
                    return "";
                }

                var row = new PracownikImportRow();

                if (isLegacy38)
                {
                    // Układ archiwalny (legacy 38 kolumn)
                    row.PracownikCol = GetCol(0);
                    row.DzialNazwa = GetCol(1);
                    row.DataRozpoczecia = GetCol(2);
                    row.DataZakonczenia = GetCol(3);
                    row.TypUmowy = GetCol(4);
                    row.JestStudentem = GetCol(5);
                    row.DataZawarcia = GetCol(6);
                    row.Imie = NormalizeForeignLetters(GetCol(8));
                    row.DrugieImie = NormalizeForeignLetters(GetCol(9));
                    row.Nazwisko = NormalizeForeignLetters(GetCol(10));
                    row.PlecExcel = GetCol(11);
                    row.BirthDateExcel = GetCol(12);
                    row.Obywatelstwo = GetCol(13);
                    row.NfzStr = GetCol(14);
                    row.UrzadSkarbowy = GetCol(15);
                    row.BankStr = GetCol(16);
                    row.Telefon = GetCol(17);
                    row.Email = GetCol(18);
                    row.Pesel = PeselHelper.NormalizePesel(GetCol(19));
                    row.Dowod = GetCol(20);
                    row.DowodWydany = GetCol(21);
                    row.DowodWazny = GetCol(22);
                    row.Paszport = GetCol(23);
                    row.PaszportWydany = GetCol(24);
                    row.PaszportWazny = GetCol(25);
                    row.PaszportBio = GetCol(26);
                    row.PaszportBioWydany = GetCol(27);
                    row.PaszportBioWazny = GetCol(28);
                    row.Ulica = GetCol(29);
                    string rawH = GetCol(30);
                    var (dh, lh) = AddressNormalizer.SplitHouseAndApartment(rawH);
                    row.NrDomu = dh;
                    row.NrLokalu = lh;
                    row.Miejscowosc = GetCol(31);
                    var (nk, _) = AddressNormalizer.NormalizePostalCode(GetCol(32), GetCol(36));
                    row.KodPocztowy = nk;
                    row.Gmina = GetCol(33);
                    row.Powiat = GetCol(34);
                    row.Wojewodztwo = GetCol(35);
                    row.Kraj = GetCol(36);
                    row.KodZawodu = GetCol(37);
                    row.KodTytuluUbezpieczenia = GetCol(38);
                }
                else
                {
                    // =========================================================================
                    // NOWY STANDARDOWY UKŁAD STAŁY (36 kolumn - zgodnie z plikiem wzorcowym):
                    // =========================================================================
                    // Col 0:  Nazwisko i imię
                    row.PracownikCol = GetCol(0);

                    // Col 1:  Fabryka (Dział)
                    row.DzialNazwa = GetCol(1);

                    // Col 2:  Umowa od
                    row.DataRozpoczecia = GetCol(2);

                    // Col 3:  Umowa do
                    row.DataZakonczenia = GetCol(3);

                    // Col 4:  Student do 26
                    row.JestStudentem = GetCol(4);

                    // Col 5:  Data zawarcia umowy
                    row.DataZawarcia = GetCol(5);

                    // Col 6:  Imię
                    row.Imie = NormalizeForeignLetters(GetCol(6));

                    // Col 7:  Drugie imię
                    row.DrugieImie = NormalizeForeignLetters(GetCol(7));

                    // Col 8:  Nazwisko
                    row.Nazwisko = NormalizeForeignLetters(GetCol(8));

                    // Col 9:  Płeć
                    row.PlecExcel = GetCol(9);

                    // Col 10: Data urodzenia
                    row.BirthDateExcel = GetCol(10);

                    // Col 11: Obywatelstwo
                    row.Obywatelstwo = GetCol(11);

                    // Col 12: Oddział NFZ
                    row.NfzStr = GetCol(12);

                    // Col 13: Urząd skarbowy
                    row.UrzadSkarbowy = GetCol(13);

                    // Col 14: Rachunek bankowy
                    row.BankStr = GetCol(14);

                    // Col 15: Bank (nazwa instytucji podana w pliku)
                    string fileBank = GetCol(15);

                    // Col 16: Telefon
                    row.Telefon = GetCol(16);

                    // Col 17: E-mail
                    row.Email = GetCol(17);

                    // Col 18: PESEL
                    row.Pesel = PeselHelper.NormalizePesel(GetCol(18));

                    // Col 19: Dowód osobisty (PL) — nr
                    row.Dowod = GetCol(19);

                    // Col 20: Dowód osobisty (PL) — wydany
                    row.DowodWydany = GetCol(20);

                    // Col 21: Dowód osobisty (PL) — ważny do
                    row.DowodWazny = GetCol(21);

                    // Col 22..24: Dowód tożsamości UE/EOG (opcjonalny / uzupełniający)
                    if (string.IsNullOrWhiteSpace(row.Dowod))
                    {
                        string dowodUe = GetCol(22);
                        if (!string.IsNullOrWhiteSpace(dowodUe))
                        {
                            row.Dowod = dowodUe;
                            if (string.IsNullOrWhiteSpace(row.DowodWydany)) row.DowodWydany = GetCol(23);
                            if (string.IsNullOrWhiteSpace(row.DowodWazny)) row.DowodWazny = GetCol(24);
                        }
                    }

                    // Col 25: Paszport — nr
                    row.Paszport = GetCol(25);

                    // Col 26: Paszport — wydany
                    row.PaszportWydany = GetCol(26);

                    // Col 27: Paszport — ważny do
                    row.PaszportWazny = GetCol(27);

                    // Col 28: Ulica
                    row.Ulica = GetCol(28);

                    // Col 29: Numer domu (z ewentualnym numerem lokalu np. 15/4)
                    string rawHouse = GetCol(29);
                    var (nrDomu, nrLokalu) = AddressNormalizer.SplitHouseAndApartment(rawHouse);
                    row.NrDomu = nrDomu;
                    row.NrLokalu = nrLokalu;

                    // Col 30: Miejscowość
                    row.Miejscowosc = GetCol(30);

                    // Col 31: Kod pocztowy
                    string rawPostcode = GetCol(31);
                    var (normKod, _) = AddressNormalizer.NormalizePostalCode(rawPostcode, "Polska");
                    row.KodPocztowy = normKod;
                    row.Kraj = "Polska";

                    // Col 32: Gmina
                    row.Gmina = GetCol(32);

                    // Col 33: Powiat
                    row.Powiat = GetCol(33);

                    // Col 34: Województwo
                    row.Wojewodztwo = GetCol(34);

                    // Col 35: Kod zawodu
                    row.KodZawodu = GetCol(35);

                    // Opcjonalne dodatkowe kolumny (jeśli występują):
                    // Col 36: Kod tytułu ubezpieczenia ZUS
                    row.KodTytuluUbezpieczenia = GetCol(36);

                    // Col 37: Stawka
                    row.Stawka = GetCol(37);

                    // Rozpoznawanie nazwy banku
                    if (!string.IsNullOrWhiteSpace(row.BankStr))
                    {
                        if (ParseBankString(row.BankStr, out string iban, out _, out string detectedBankName))
                        {
                            row.BankNazwa = (!string.IsNullOrWhiteSpace(detectedBankName) && detectedBankName != "Rachunek podstawowy")
                                ? detectedBankName
                                : (!string.IsNullOrWhiteSpace(fileBank) ? fileBank : detectedBankName);

                            if (!BankClearingResolver.ValidateIbanChecksum(iban))
                            {
                                row.Ostrzezenia += "Niepoprawny numer rachunku bankowego (IBAN / suma kontrolna). ";
                            }
                        }
                        else
                        {
                            row.BankNazwa = !string.IsNullOrWhiteSpace(fileBank) ? fileBank : "Rachunek podstawowy";
                            row.Ostrzezenia += "Niepoprawny numer rachunku bankowego (IBAN). ";
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(fileBank))
                    {
                        row.BankNazwa = fileBank;
                    }
                }

                // Fallback dla imienia i nazwiska z pełnej kolumny, jeśli osobne pola były puste
                if (string.IsNullOrWhiteSpace(row.Imie) && string.IsNullOrWhiteSpace(row.Nazwisko) && !string.IsNullOrWhiteSpace(row.PracownikCol))
                {
                    var parts = row.PracownikCol.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2)
                    {
                        row.Nazwisko = parts[0];
                        row.Imie = string.Join(" ", parts.Skip(1));
                    }
                    else if (parts.Length == 1)
                    {
                        row.Nazwisko = parts[0];
                    }
                }

                // Normalizacja i rozdzielenie numeru domu oraz lokalu z ulicy (np. 'Chodzki 5' + '60' -> 'Chodzki', dom='5', lok='60')
                if (!string.IsNullOrWhiteSpace(row.Ulica))
                {
                    var (cleanSt, extDom, extLok) = AddressNormalizer.ExtractHouseFromStreetIfMissing(row.Ulica, row.NrDomu, row.NrLokalu);
                    row.Ulica = cleanSt;
                    row.NrDomu = extDom;
                    if (!string.IsNullOrWhiteSpace(extLok)) row.NrLokalu = extLok;
                }

                // Bezpieczne autouzupełnianie brakujących danych adresowych z bazy nexo (DaneAdresowe.db)
                AddressEnrichmentService.Instance.EnrichRowAddress(row);

                // Address sanitation
                SanitizeAddressFields(row);

                // Walidacja obecności podstawowych danych pracownika
                if (!string.IsNullOrWhiteSpace(row.Imie) ||
                    !string.IsNullOrWhiteSpace(row.Nazwisko) ||
                    !string.IsNullOrWhiteSpace(row.Pesel) ||
                    !string.IsNullOrWhiteSpace(row.Paszport) ||
                    !string.IsNullOrWhiteSpace(row.Dowod) ||
                    !string.IsNullOrWhiteSpace(row.PaszportBio))
                {
                    rows.Add(row);
                }
            }

            return rows;
        }

        private static List<IReadOnlyList<string>> ReadExcelRows(string filePath)
        {
            var result = new List<IReadOnlyList<string>>();
            var queryRows = MiniExcel.Query(filePath, useHeaderRow: false, excelType: ExcelType.XLSX)
                                     .Cast<IDictionary<string, object>>();

            foreach (var dict in queryRows)
            {
                var rowList = new List<string>();
                if (dict != null)
                {
                    foreach (var val in dict.Values)
                    {
                        rowList.Add(val?.ToString()?.Trim() ?? "");
                    }
                }
                result.Add(rowList);
            }
            return result;
        }

        public static void SanitizeAddressFields(PracownikImportRow row)
        {
            if (row == null) return;
            string imie = (row.Imie ?? "").Trim();
            string nazwisko = (row.Nazwisko ?? "").Trim();

            if (!string.IsNullOrWhiteSpace(row.Ulica))
            {
                string ulicaTrim = row.Ulica.Trim();
                bool isPib = (!string.IsNullOrEmpty(imie) && ulicaTrim.IndexOf(imie, StringComparison.OrdinalIgnoreCase) >= 0) ||
                             (!string.IsNullOrEmpty(nazwisko) && ulicaTrim.IndexOf(nazwisko, StringComparison.OrdinalIgnoreCase) >= 0);
                bool isPhone = Regex.IsMatch(ulicaTrim, @"^\+?\d[\d\s\-]{7,}\d$");

                if (isPib || isPhone)
                {
                    row.Ulica = "";
                    row.Ostrzezenia += $"Wykryto błędną treść w polu Ulica ({(isPib ? "Imię/Nazwisko" : "Numer telefonu")}). Pole zostało oczyszczone. ";
                }
                else if (row.Ulica.Length > ZusSchemaLimits.ZusUlicaMaxLen)
                {
                    row.Ulica = row.Ulica.Substring(0, ZusSchemaLimits.ZusUlicaMaxLen).Trim();
                }
            }

            if (!string.IsNullOrWhiteSpace(row.Miejscowosc) && row.Miejscowosc.Length > ZusSchemaLimits.ZusMiejscowoscMaxLen)
            {
                row.Miejscowosc = row.Miejscowosc.Substring(0, ZusSchemaLimits.ZusMiejscowoscMaxLen).Trim();
            }
        }

        private static bool ParseBankString(string rawBankStr, out string iban, out string currencySymbol, out string bankName)
        {
            iban = "";
            currencySymbol = "PLN";
            bankName = "Rachunek podstawowy";

            if (string.IsNullOrWhiteSpace(rawBankStr)) return false;

            var match = Regex.Match(rawBankStr, @"(PL\s*)?\d{2}(\s*\d{4}){6}|\d{26}", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                iban = match.Value.Replace(" ", "").ToUpper();
                if (iban.Length > 34) iban = iban.Substring(0, 34);

                string remainder = (rawBankStr.Substring(0, match.Index) + rawBankStr.Substring(match.Index + match.Length)).Trim();
                remainder = remainder.Trim('\r', '\n', ' ', '-', ':', '|');

                var currMatch = Regex.Match(remainder, @"^(PLN|EUR|USD|GBP|CHF)", RegexOptions.IgnoreCase);
                if (currMatch.Success) currencySymbol = currMatch.Value.ToUpper();
                bankName = BankClearingResolver.GetBankName(iban, "Rachunek podstawowy");
                return true;
            }

            string[] bankParts = rawBankStr.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            iban = bankParts.Length >= 1 ? bankParts[0].Trim().Replace(" ", "").ToUpper() : "";
            if (iban.Length > 34) iban = iban.Substring(0, 34);
            bankName = BankClearingResolver.GetBankName(iban, "Rachunek podstawowy");
            return !string.IsNullOrWhiteSpace(iban);
        }

        private static string NormalizeForeignLetters(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            return input.Trim();
        }
    }
}
