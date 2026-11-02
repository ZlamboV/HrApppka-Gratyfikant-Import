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
    /// Uniwersalny loader plików importu pracowników obsługujący pliki .xlsx (poprzez MiniExcel)
    /// oraz pliki .csv (poprzez streamingowy parser Rfc4180CsvReader).
    /// Zwraca zmapowane i wstępnie znormalizowane obiekty PracownikImportRow.
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
            if (headerRow.Count < 30)
            {
                throw new InvalidOperationException($"Plik ma nieprawidłową liczbę kolumn ({headerRow.Count}). Oczekiwano co najmniej 30 kolumn (standard HrAppka to 38 kolumn).");
            }

            var rows = new List<PracownikImportRow>();

            for (int r = 1; r < rawRows.Count; r++)
            {
                var values = rawRows[r];

                string GetColVal(int idx1Based)
                {
                    int zeroIdx = idx1Based - 1;
                    if (zeroIdx >= 0 && zeroIdx < values.Count)
                    {
                        return values[zeroIdx]?.Trim() ?? "";
                    }
                    return "";
                }

                var row = new PracownikImportRow();

                // Col 1: Full name
                row.PracownikCol = GetColVal(1);

                // Col 2-7: Employment & Contract
                row.DzialNazwa = GetColVal(2);
                row.DataRozpoczecia = GetColVal(3);
                row.DataZakonczenia = GetColVal(4);
                row.TypUmowy = GetColVal(5);
                row.JestStudentem = GetColVal(6);
                row.DataZawarcia = GetColVal(7);
                // Col 8: Status (ignored)

                // Col 9-11: Names
                row.Imie = NormalizeForeignLetters(GetColVal(9));
                row.DrugieImie = NormalizeForeignLetters(GetColVal(10));
                row.Nazwisko = NormalizeForeignLetters(GetColVal(11));

                // Col 12-19: Demographics & Contact
                row.PlecExcel = GetColVal(12);
                row.BirthDateExcel = GetColVal(13);
                row.Obywatelstwo = GetColVal(14);
                row.NfzStr = GetColVal(15);
                row.UrzadSkarbowy = GetColVal(16);
                row.BankStr = GetColVal(17);
                row.Telefon = GetColVal(18);
                row.Email = GetColVal(19);

                // Col 20: PESEL
                row.Pesel = PeselHelper.NormalizePesel(GetColVal(20));

                // Col 21-23: Dowód osobisty
                row.Dowod = GetColVal(21);
                row.DowodWydany = GetColVal(22);
                row.DowodWazny = GetColVal(23);

                // Col 24-26: Paszport
                row.Paszport = GetColVal(24);
                row.PaszportWydany = GetColVal(25);
                row.PaszportWazny = GetColVal(26);

                // Col 27-29: Paszport biometryczny
                row.PaszportBio = GetColVal(27);
                row.PaszportBioWydany = GetColVal(28);
                row.PaszportBioWazny = GetColVal(29);

                // Col 30-38: Address & Occupation
                row.Ulica = GetColVal(30);
                string rawHouse = GetColVal(31);
                var (nrDomu, nrLokalu) = AddressNormalizer.SplitHouseAndApartment(rawHouse);
                row.NrDomu = nrDomu;
                row.NrLokalu = nrLokalu;

                row.Miejscowosc = GetColVal(32);
                var (normKod, _) = AddressNormalizer.NormalizePostalCode(GetColVal(33), GetColVal(37));
                row.KodPocztowy = normKod;
                row.Gmina = GetColVal(34);
                row.Powiat = GetColVal(35);
                row.Wojewodztwo = GetColVal(36);
                row.Kraj = GetColVal(37);
                row.KodZawodu = GetColVal(38);

                // Address sanitation
                SanitizeAddressFields(row);

                // Bank account format check
                if (!string.IsNullOrWhiteSpace(row.BankStr))
                {
                    if (!ParseBankString(row.BankStr, out string iban, out _, out _) || string.IsNullOrWhiteSpace(iban) || iban.Length < 10 || iban.Length > 34)
                    {
                        row.Ostrzezenia += "Niepoprawny numer rachunku bankowego (IBAN). ";
                    }
                }

                // Filter empty row
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
                return true;
            }

            string[] bankParts = rawBankStr.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            iban = bankParts.Length >= 1 ? bankParts[0].Trim().Replace(" ", "").ToUpper() : "";
            if (iban.Length > 34) iban = iban.Substring(0, 34);
            return !string.IsNullOrWhiteSpace(iban);
        }

        private static string NormalizeForeignLetters(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            return input.Trim();
        }
    }
}
