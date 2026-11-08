using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace HrAppka_Import_Pracowników.Services
{
    /// <summary>
    /// Rozpoznaje oficjalną nazwę banku w Polsce na podstawie 8-cyfrowego numeru rozliczeniowego
    /// jednostki organizacyjnej banku (NRB / IBAN), zgodnie z ewidencją NBP / KNF.
    /// </summary>
    public static class BankClearingResolver
    {
        private static readonly Dictionary<string, string> BankInstitutions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "1010", "Narodowy Bank Polski" },
            { "1020", "Powszechna Kasa Oszczędności Bank Polski SA" },
            { "1030", "Bank Handlowy w Warszawie SA (Citi Handlowy)" },
            { "1050", "ING Bank Śląski SA" },
            { "1060", "Bank BPH SA" },
            { "1090", "Santander Bank Polska SA" },
            { "1130", "Bank Gospodarstwa Krajowego" },
            { "1140", "mBank SA" },
            { "1160", "Bank Millennium SA" },
            { "1220", "Plus Bank SA" },
            { "1240", "Bank Polska Kasa Opieki SA (Pekao)" },
            { "1300", "Bank Ochrony Środowiska SA" },
            { "1320", "Bank Pocztowy SA" },
            { "1440", "Powszechna Kasa Oszczędności Bank Polski SA (Nordea)" },
            { "1470", "Euro Bank SA (Millennium)" },
            { "1500", "Santander Consumer Bank SA" },
            { "1540", "Bank Ochrony Środowiska SA (BOŚ)" },
            { "1580", "Mercedes-Benz Bank Polska SA" },
            { "1600", "BNP Paribas Bank Polska SA" },
            { "1610", "SGB-Bank SA" },
            { "1670", "RBS Bank (Polska) SA" },
            { "1680", "Plus Bank SA" },
            { "1750", "Raiffeisen Bank International AG" },
            { "1840", "Société Générale SA Oddział w Polsce" },
            { "1870", "Nest Bank SA" },
            { "1910", "Deutsche Bank Polska SA" },
            { "1930", "Bank Polskiej Spółdzielczości SA (BPS)" },
            { "1940", "Credit Agricole Bank Polska SA" },
            { "1950", "Idea Bank SA" },
            { "2000", "Rabobank Polska SA" },
            { "2030", "BNP Paribas Bank Polska SA (BGŻ)" },
            { "2120", "Santander Consumer Bank SA" },
            { "2130", "Volkswagen Bank GmbH Oddział w Polsce" },
            { "2140", "FCA Bank SpA Oddział w Polsce" },
            { "2160", "Toyota Bank Polska SA" },
            { "2190", "DNB Bank Polska SA" },
            { "2480", "Alior Bank SA" },
            { "2490", "Alior Bank SA" }
        };

        /// <summary>
        /// Czyści numer konta ze spacji, myślników i prefiksu PL.
        /// </summary>
        public static string NormalizeNrb(string? rawAccount)
        {
            if (string.IsNullOrWhiteSpace(rawAccount)) return string.Empty;
            string cleaned = Regex.Replace(rawAccount, @"[\s\-]", "").ToUpperInvariant();
            if (cleaned.StartsWith("PL", StringComparison.OrdinalIgnoreCase))
            {
                cleaned = cleaned.Substring(2);
            }
            return cleaned;
        }

        /// <summary>
        /// Wyciąga do 8 cyfr numeru rozliczeniowego (pozycje od 3 cyfry w NRB).
        /// Działa nawet w przypadku literówek w długości rachunku (np. 27 cyfr).
        /// </summary>
        public static string? ExtractSortCode(string? rawAccount)
        {
            string nrb = NormalizeNrb(rawAccount);
            if (nrb.Length >= 6 && Regex.IsMatch(nrb, @"^\d+$"))
            {
                int len = Math.Min(8, nrb.Length - 2);
                return nrb.Substring(2, len);
            }
            return null;
        }

        /// <summary>
        /// Alias dla ExtractSortCode.
        /// </summary>
        public static string? ExtractClearingNumber(string? rawAccount) => ExtractSortCode(rawAccount);

        /// <summary>
        /// Formatuje numer konta do postaci ze spacjami (KK BBBB BBBB BBBB BBBB BBBB BBBB).
        /// </summary>
        public static string FormatNrb(string? rawAccount)
        {
            string nrb = NormalizeNrb(rawAccount);
            if (string.IsNullOrWhiteSpace(nrb)) return string.Empty;
            if (nrb.Length < 6) return nrb;

            var parts = new List<string> { nrb.Substring(0, 2) };
            int pos = 2;
            while (pos < nrb.Length)
            {
                int take = Math.Min(4, nrb.Length - pos);
                parts.Add(nrb.Substring(pos, take));
                pos += take;
            }
            return string.Join(" ", parts);
        }

        /// <summary>
        /// Rozpoznaje oficjalną nazwę banku na podstawie numeru konta.
        /// Zwraca dopasowaną nazwę lub domyślny fallback.
        /// </summary>
        public static string GetBankName(string? rawAccount, string fallback = "Rachunek podstawowy")
        {
            string? sortCode = ExtractSortCode(rawAccount);
            if (sortCode != null)
            {
                // Pierwsze 4 cyfry
                string inst4 = sortCode.Substring(0, 4);
                if (BankInstitutions.TryGetValue(inst4, out string? name4))
                {
                    return name4;
                }

                // Banki spółdzielcze zaczynające się od 8 lub 9
                char firstDigit = sortCode[0];
                if (firstDigit == '8')
                {
                    return "Bank Spółdzielczy (zrzeszenie SGB)";
                }
                if (firstDigit == '9')
                {
                    return "Bank Spółdzielczy (zrzeszenie BPS)";
                }
            }

            return fallback;
        }

        /// <summary>
        /// Waliduje sumę kontrolną polskiego numeru konta (NRB / IBAN) algorytmem Modulo 97.
        /// </summary>
        public static bool ValidateIbanChecksum(string? rawAccount)
        {
            if (string.IsNullOrWhiteSpace(rawAccount)) return false;
            string nrb = NormalizeNrb(rawAccount);
            if (nrb.Length != 26 || !Regex.IsMatch(nrb, @"^\d{26}$"))
            {
                // Jeśli to zagraniczny IBAN (np. UA, DE): min 15, max 34 znaki alfanumeryczne
                string rawUpper = Regex.Replace(rawAccount, @"[\s\-]", "").ToUpperInvariant();
                return rawUpper.Length >= 15 && rawUpper.Length <= 34 && Regex.IsMatch(rawUpper, @"^[A-Z]{2}[A-Z0-9]+$");
            }

            // Algorytm Modulo 97 dla polskiego NRB:
            // Przeniesienie cyfr kontrolnych na koniec + dopisanie PL (P=25, L=21)
            // NRB: KK BBBBBBBB NNNNNNNNNNNNNNNN -> BBBBBBBB NNNNNNNNNNNNNNNN 25 21 KK
            string rearranged = nrb.Substring(2) + "2521" + nrb.Substring(0, 2);

            int remainder = 0;
            foreach (char c in rearranged)
            {
                int digit = c - '0';
                remainder = (remainder * 10 + digit) % 97;
            }

            return remainder == 1;
        }
    }
}
