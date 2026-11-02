using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace HrAppka_Import_Pracowników.Services
{
    /// <summary>
    /// Autorytatywne limity długości pól wynikające ze schematu ZUS KEDU 5.7 (Platnik_5_7.xsd) oraz bazy Nexo.
    /// </summary>
    public static class ZusSchemaLimits
    {
        // Limity ZUS Płatnik KEDU 5.7 XSD
        public const int ZusUlicaMaxLen = 30;         // t_Ulica_30
        public const int ZusNrDomuMaxLen = 7;         // t_NrLokaluDomu_7
        public const int ZusNrLokaluMaxLen = 7;       // t_NrLokaluDomu_7
        public const int ZusMiejscowoscMaxLen = 26;   // t_Miejscowosc_26
        public const int ZusPocztaMaxLen = 26;        // t_Miejscowosc_26
        public const int ZusGminaMaxLen = 26;         // t_Gmina_26
        public const int ZusKodPocztowyDigits = 5;    // t_KodPocztowy_5
        public const int ZusSeriaNrDokMaxLen = 9;     // t_SeriaNrDok
        public const int ZusImieMaxLen = 22;          // t_Imie_22
        public const int ZusNazwiskoMaxLen = 31;      // t_Nazwisko_31
        public const int ZusTelefonMaxLen = 12;       // t_Telefon_12
        public const int ZusEmailMaxLen = 30;         // t_eMail_30

        // Limity Bazy Danych Nexo (Dokumentacja_bazy_danych_nexo.htm)
        public const int NexoDbUlicaMaxLen = 64;      // AdresySzczegoly.Ulica nvarchar(64)
        public const int NexoDbNrDomuMaxLen = 10;     // AdresySzczegoly.NrDomu nvarchar(10)
        public const int NexoDbNrLokaluMaxLen = 10;   // AdresySzczegoly.NrLokalu nvarchar(10)
        public const int NexoDbMiejscowoscMaxLen = 64;// AdresySzczegoly.Miejscowosc nvarchar(64)
        public const int NexoDbPocztaMaxLen = 64;     // AdresySzczegoly.Poczta nvarchar(64)
        public const int NexoDbKodPocztowyMaxLen = 10;// AdresySzczegoly.KodPocztowy nvarchar(10)
        public const int NexoDbSeriaNrDokMaxLen = 20; // Osoby.SeriaINumerDokumentuTozsamosci nvarchar(20)
    }

    /// <summary>
    /// Reprezentuje w pełni znormalizowane i przycięte dane adresowe zgodne ze schematem ZUS KEDU oraz bazą Nexo.
    /// </summary>
    public sealed class NormalizedAddress
    {
        public string Ulica { get; init; } = string.Empty;
        public string NrDomu { get; init; } = "1";
        public string NrLokalu { get; init; } = string.Empty;
        public string Miejscowosc { get; init; } = string.Empty;
        public string Poczta { get; init; } = string.Empty;
        public string KodPocztowy { get; init; } = string.Empty;
        public string ZusKodPocztowy { get; init; } = string.Empty;
        public string Gmina { get; init; } = string.Empty;
        public string Powiat { get; init; } = string.Empty;
        public string Wojewodztwo { get; init; } = string.Empty;
        public string Kraj { get; init; } = "Polska";
        public bool IsValid => !string.IsNullOrWhiteSpace(Miejscowosc) && !string.IsNullOrWhiteSpace(NrDomu);
        public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
    }

    /// <summary>
    /// Kompleksowy silnik normalizacji danych adresowych, rozdzielania numerów domów/lokali
    /// oraz rygorystycznego wymuszania limitów ZUS KEDU i bazy danych Nexo.
    /// </summary>
    public static class AddressNormalizer
    {
        private static readonly Regex ControlCharsRegex = new Regex(@"[\x00-\x1F\x7F]", RegexOptions.Compiled);
        private static readonly Regex MultiWhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);

        // Regex 1: Złożony numer domu i lokalu (np. "4/64", "12A m. 5", "15B/2", "15 lokal 3", "24/5A")
        private static readonly Regex CompoundHouseAptRegex = new Regex(
            @"^([0-9]+[A-Za-z]?|[A-Za-z]+[0-9]+)\s*(?:[\/\\,]|(?:\s+m\.?|\s+lok\.?|\s+lokal|\s+apt\.?|\s+mieszkanie|\s+nr))\s*(?:m\.?|lok\.?|lokal|apt\.?|mieszkanie|nr)?\s*([0-9]+[A-Za-z]?|[A-Za-z])$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase
        );

        // Regex 2: Samotny znacznik lokalu w polu domu (np. "m. 5", "lok. 12")
        private static readonly Regex StandaloneApartmentRegex = new Regex(
            @"^(?:m\.?|lok\.?|lokal|apt\.?|mieszkanie)\s*([0-9]+[A-Za-z]?)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase
        );

        // Regex 3: Wyciąganie numeru domu i lokalu z końca nazwy ulicy (np. "Kwiatowa 12/4", "Lipowa 12 m. 4", "Marszałkowska 12A m. 5")
        private static readonly Regex StreetTrailingNumberRegex = new Regex(
            @"^(.*?)\s+([0-9]+[A-Za-z]?(?:\s*(?:[\/\\,]|m\.|m\b|lok\.|lok\b|lokal\b|apt\.|apt\b|mieszkanie\b|nr\.|nr\b)\s*.*)?)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase
        );

        // Regex 4: Usuwanie zbędnego prefiksu ulicy ("ul.", "ulica ")
        private static readonly Regex StreetPrefixRegex = new Regex(
            @"^(?:ul\.\s*|ulica\s+|ul\s+)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase
        );

        // Regexy kodu pocztowego
        private static readonly Regex PostalCodeCleanRegex = new Regex(@"[^\w\-]", RegexOptions.Compiled);
        private static readonly Regex PolishPostalPrefixRegex = new Regex(@"^(?:PL)[\s\-]*", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex FormattedPostalRegex = new Regex(@"^\d{2}-\d{3}$", RegexOptions.Compiled);

        public static string CleanWhitespace(string? input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            string noCtrl = ControlCharsRegex.Replace(input, string.Empty);
            return MultiWhitespaceRegex.Replace(noCtrl, " ").Trim();
        }

        public static string CleanAndTruncate(string? input, int maxLength, bool trimEndDelimiters = false)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            string cleaned = CleanWhitespace(input);
            if (trimEndDelimiters)
            {
                cleaned = cleaned.TrimEnd('/', '\\', ',', '.', '-', ' ');
            }
            if (cleaned.Length <= maxLength)
                return cleaned;

            return cleaned.Substring(0, maxLength).TrimEnd();
        }

        public static NormalizedAddress Normalize(
            string? rawStreet,
            string? rawHouseAndApt,
            string? rawCity,
            string? rawPostalCode,
            string? rawGmina = null,
            string? rawPowiat = null,
            string? rawWojewodztwo = null,
            string? rawCountry = null)
        {
            var warnings = new List<string>();

            string city = CleanWhitespace(rawCity);
            string street = CleanWhitespace(rawStreet);
            string houseInput = CleanWhitespace(rawHouseAndApt);
            string country = string.IsNullOrWhiteSpace(rawCountry) ? "Polska" : CleanWhitespace(rawCountry);

            // 1. Jeśli ulica kończy się nazwą miejscowości (np. "Kwiatowa, Lublin"), usuń powtórzoną miejscowość
            if (!string.IsNullOrWhiteSpace(city) && street.EndsWith(city, StringComparison.OrdinalIgnoreCase))
            {
                street = street.Substring(0, street.Length - city.Length).TrimEnd(',', ' ', '-');
            }

            // 2. Usuń zbędny prefiks "ul." / "ulica "
            string streetNoPrefix = StreetPrefixRegex.Replace(street, string.Empty).Trim();
            if (streetNoPrefix.Length > 0)
            {
                street = streetNoPrefix;
            }

            // 3. Rozdziel numer domu i lokalu z pola houseInput
            var (nrDomu, nrLokalu) = SplitHouseAndApartment(houseInput);

            // 4. Jeśli w polu domu nie było numeru domu, sprawdź czy znajduje się na końcu nazwy ulicy (np. "Kwiatowa 12/4")
            var (cleanStreet, extractedDom, extractedLok) = ExtractHouseFromStreetIfMissing(street, nrDomu, nrLokalu);
            street = cleanStreet;
            if (string.IsNullOrWhiteSpace(nrDomu) && !string.IsNullOrWhiteSpace(extractedDom))
            {
                nrDomu = extractedDom;
                if (string.IsNullOrWhiteSpace(nrLokalu) && !string.IsNullOrWhiteSpace(extractedLok))
                {
                    nrLokalu = extractedLok;
                }
            }

            // 5. Wymóg Gratyfikant Nexo: NrDomu nie może być pusty. W przypadku braku ustaw domyślnie "1"
            if (string.IsNullOrWhiteSpace(nrDomu))
            {
                nrDomu = "1";
                warnings.Add("Brak numeru domu - ustawiono domyślnie '1'.");
            }

            // 6. Normalizacja kodu pocztowego (format "XX-XXX" dla Nexo oraz 5 cyfr dla ZUS)
            var (formattedKod, zusKod) = NormalizePostalCode(rawPostalCode, country);
            if (string.IsNullOrWhiteSpace(formattedKod))
            {
                warnings.Add("Brak poprawnego kodu pocztowego.");
            }

            // 7. Rygorystyczne przycinanie do limitów ZUS i Nexo
            // Limit ZUS dla ulicy to 30; limit bazy Nexo to 64. Przycinając do 30 znaków zapobiegamy zarówno błędom ZUS jak i wyjątkom SQL
            string finalStreet = CleanAndTruncate(street, ZusSchemaLimits.ZusUlicaMaxLen);
            if (street.Length > ZusSchemaLimits.ZusUlicaMaxLen)
            {
                warnings.Add($"Ulica '{street}' została skrócona do {ZusSchemaLimits.ZusUlicaMaxLen} znaków: '{finalStreet}'.");
            }

            string finalNrDomu = CleanAndTruncate(nrDomu, ZusSchemaLimits.ZusNrDomuMaxLen, trimEndDelimiters: true);
            string finalNrLokalu = CleanAndTruncate(nrLokalu, ZusSchemaLimits.ZusNrLokaluMaxLen, trimEndDelimiters: true);

            string finalCity = CleanAndTruncate(city, ZusSchemaLimits.ZusMiejscowoscMaxLen);
            if (city.Length > ZusSchemaLimits.ZusMiejscowoscMaxLen)
            {
                warnings.Add($"Miejscowość '{city}' została skrócona do {ZusSchemaLimits.ZusMiejscowoscMaxLen} znaków: '{finalCity}'.");
            }
            if (string.IsNullOrWhiteSpace(finalCity))
            {
                warnings.Add("Brak miejscowości.");
            }

            string finalPoczta = finalCity;

            string gmina = CleanWhitespace(rawGmina);
            string finalGmina = CleanAndTruncate(gmina, ZusSchemaLimits.ZusGminaMaxLen);

            return new NormalizedAddress
            {
                Ulica = finalStreet,
                NrDomu = finalNrDomu,
                NrLokalu = finalNrLokalu,
                Miejscowosc = finalCity,
                Poczta = finalPoczta,
                KodPocztowy = formattedKod,
                ZusKodPocztowy = zusKod,
                Gmina = finalGmina,
                Powiat = CleanWhitespace(rawPowiat),
                Wojewodztwo = CleanWhitespace(rawWojewodztwo),
                Kraj = country,
                Warnings = warnings
            };
        }

        /// <summary>
        /// Normalizuje jedno-liniowy adres (np. "Marszałkowska 12A m. 5"), w którym ulica, numer domu i lokalu
        /// znajdują się w jednym ciągu tekstowym.
        /// </summary>
        public static NormalizedAddress NormalizeSingleAddressLine(string? rawAddress, string? rawCity, string? rawPostalCode, string? rawCountry = null)
        {
            return Normalize(rawAddress, null, rawCity, rawPostalCode, null, null, null, rawCountry);
        }

        /// <summary>
        /// Przeciążenie zgodne z kontraktem interfejsowym w PROJECT.md:
        /// AddressNormalizer.Normalize(string rawAddress, string rawMiejscowosc, string rawKod) -> NormalizedAddress
        /// Posiada dokładnie 3 parametry pozycyjne bez wartości domyślnych, co zapobiega kolizji przeciążeń.
        /// </summary>
        public static NormalizedAddress Normalize(string? rawAddress, string? rawCity, string? rawPostalCode)
        {
            return NormalizeSingleAddressLine(rawAddress, rawCity, rawPostalCode, null);
        }

        public static (string nrDomu, string nrLokalu) SplitHouseAndApartment(string? rawHouse)
        {
            if (string.IsNullOrWhiteSpace(rawHouse))
                return (string.Empty, string.Empty);

            string s = rawHouse.Trim().TrimEnd('/', '\\', ',', ' ');

            // Dopasowanie 1: Złożony numer ("4/64", "12A m. 5", "15B/2", "15 lokal 3", "24/5A")
            var match = CompoundHouseAptRegex.Match(s);
            if (match.Success)
            {
                string dom = match.Groups[1].Value.Trim();
                string lok = match.Groups[2].Value.Trim();
                return (CleanAndTruncate(dom, ZusSchemaLimits.ZusNrDomuMaxLen, true), 
                        CleanAndTruncate(lok, ZusSchemaLimits.ZusNrLokaluMaxLen, true));
            }

            // Dopasowanie 2: Dzielenie po ukośniku (fallback dla niestandardowych formatów)
            if (s.Contains('/') || s.Contains('\\'))
            {
                var parts = s.Split(new[] { '/', '\\' }, 2);
                string dom = parts[0].Trim();
                string lok = parts[1].Trim().TrimStart('m', 'M', '.', 'l', 'L', 'o', 'k', 'K', 'a', 'A', 'p', 'P', 't', 'T', 'n', 'N', 'r', 'R', ' ');

                var subMatch = CompoundHouseAptRegex.Match(dom);
                if (subMatch.Success)
                {
                    dom = subMatch.Groups[1].Value.Trim();
                    if (string.IsNullOrWhiteSpace(lok))
                    {
                        lok = subMatch.Groups[2].Value.Trim();
                    }
                }

                return (CleanAndTruncate(dom, ZusSchemaLimits.ZusNrDomuMaxLen, true), 
                        CleanAndTruncate(lok, ZusSchemaLimits.ZusNrLokaluMaxLen, true));
            }

            // Dopasowanie 3: Samotny numer lokalu (np. "m. 5")
            var matchApt = StandaloneApartmentRegex.Match(s);
            if (matchApt.Success)
            {
                return (string.Empty, CleanAndTruncate(matchApt.Groups[1].Value.Trim(), ZusSchemaLimits.ZusNrLokaluMaxLen, true));
            }

            // W przeciwnym razie: tylko numer domu
            return (CleanAndTruncate(s, ZusSchemaLimits.ZusNrDomuMaxLen, true), string.Empty);
        }

        public static (string Formatted, string ZusDigits) NormalizePostalCode(string? rawCode, string? countryCode = "PL")
        {
            if (string.IsNullOrWhiteSpace(rawCode))
                return (string.Empty, string.Empty);

            string cleaned = PolishPostalPrefixRegex.Replace(rawCode.Trim(), string.Empty).Trim();
            cleaned = PostalCodeCleanRegex.Replace(cleaned, string.Empty);

            bool isPolish = string.IsNullOrWhiteSpace(countryCode) ||
                            countryCode.Equals("PL", StringComparison.OrdinalIgnoreCase) ||
                            countryCode.Equals("Polska", StringComparison.OrdinalIgnoreCase) ||
                            countryCode.Equals("Poland", StringComparison.OrdinalIgnoreCase);

            if (isPolish)
            {
                string digits = Regex.Replace(cleaned, @"\D", string.Empty);
                if (digits.Length == ZusSchemaLimits.ZusKodPocztowyDigits)
                {
                    string formatted = $"{digits.Substring(0, 2)}-{digits.Substring(2, 3)}";
                    return (formatted, digits);
                }

                if (FormattedPostalRegex.IsMatch(cleaned))
                {
                    string zusDigits = cleaned.Replace("-", string.Empty);
                    return (cleaned, zusDigits);
                }

                string fallback = cleaned.Length <= ZusSchemaLimits.NexoDbKodPocztowyMaxLen 
                    ? cleaned 
                    : cleaned.Substring(0, ZusSchemaLimits.NexoDbKodPocztowyMaxLen);
                return (fallback, digits.Length >= 5 ? digits.Substring(0, 5) : digits);
            }
            else
            {
                string foreign = cleaned.Length <= ZusSchemaLimits.NexoDbKodPocztowyMaxLen 
                    ? cleaned 
                    : cleaned.Substring(0, ZusSchemaLimits.NexoDbKodPocztowyMaxLen);
                return (foreign, foreign);
            }
        }

        private static (string CleanStreet, string NrDomu, string NrLokalu) ExtractHouseFromStreetIfMissing(
            string street,
            string currentDom,
            string currentLok)
        {
            var match = StreetTrailingNumberRegex.Match(street);
            if (match.Success)
            {
                string stName = match.Groups[1].Value.Trim();
                string stNumPart = match.Groups[2].Value.Trim();

                var (splitDom, splitLok) = SplitHouseAndApartment(stNumPart);

                if (string.IsNullOrWhiteSpace(currentDom))
                {
                    string finalDom = !string.IsNullOrWhiteSpace(splitDom) ? splitDom : stNumPart;
                    string finalLok = !string.IsNullOrWhiteSpace(currentLok) ? currentLok : splitLok;
                    return (stName, finalDom, finalLok);
                }
                else if (currentDom.Equals(stNumPart, StringComparison.OrdinalIgnoreCase) ||
                         (!string.IsNullOrWhiteSpace(splitDom) && currentDom.Equals(splitDom, StringComparison.OrdinalIgnoreCase)))
                {
                    return (stName, currentDom, currentLok);
                }
            }

            return (street, currentDom, currentLok);
        }
    }
}
