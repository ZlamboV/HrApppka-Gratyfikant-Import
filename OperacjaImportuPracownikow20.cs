#pragma warning disable CA1416
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using InsERT.Moria.Bank;
using InsERT.Moria.DeklaracjeZUS;
using InsERT.Moria.Flagi;
using InsERT.Moria.Kadry;
using InsERT.Moria.Kadry.Duze;
using InsERT.Moria.Klienci;
using InsERT.Moria.Listy;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Rozszerzanie.Operacje;
using InsERT.Moria.Sfera;
using InsERT.Moria.Waluty;
using Microsoft.Win32;
using MiniExcelLibs;
using HrAppka_Import_Pracowników.Services;

namespace HrAppka_Import_Pracowników
{
    /// <summary>
    /// ??????? ?????? ???????? ???? "Import 2.0 (Sfera Native)" ??? Gratyfikant Nexo Pro.
    /// ?????????? ?? ???? ???????? ??????????? Sfera Native 2.0 ? ?????? ?????????? ????? v1.0,
    /// ???????? ???????????? Unit of Work oraz ????????? ?????????? LogNexo.
    /// </summary>
    internal class OperacjaImportuPracownikow20 : OperacjaNaLiscieDanych<Pracownik, int>
    {
        // ?? ????????? ???????, ????? ????????? ????????????? ? ?????????? ???? ?????? Pracownicy
        protected override string[]? SciezkaWMenu => null;

        // ????? ?????? ???? ? Gratyfikant Nexo Pro
        public override string Nazwa => "Importuj z HrAppka";

        /// <summary>
        /// ????? ????????????/?????????? ?????? ???? ? UI Gratyfikant Nexo Pro.
        /// </summary>
        protected override bool SprawdzCzyMoznaWykonac(
            IReadOnlyCollection<int> identyfikatoryWybranychElementow,
            IKontekstListyDanych kontekstListyDanych,
            IKontekstOperacji kontekstOperacji)
        {
            return true;
        }

        /// <summary>
        /// ????? ????? ???????? ??????? 2.0.
        /// </summary>
        protected override void Wykonaj(
            IReadOnlyCollection<int> identyfikatoryWybranychElementow,
            IKontekstListyDanych kontekstListyDanych,
            IKontekstOperacji kontekstOperacji)
        {
            var uchwyt = kontekstOperacji.Uchwyt;
            LogNexo.Informacja("[Import 2.0] Rozpoczto operacj Import 2.0 (Sfera Native)");

            // 1. ????????? ??????? ?????? ????? Excel
            if (!PodajOSciezkeDoPliku(out string sciezka))
            {
                LogNexo.Informacja("[Import 2.0] Anulowano wybr pliku Excel.");
                return;
            }

            List<PracownikImportRow> rows;
            try
            {
                LogNexo.Informacja($"[Import 2.0] Wczytywanie pliku Excel: {sciezka}");
                rows = WczytajExcel(sciezka);
                LogNexo.Informacja($"[Import 2.0] Pomylnie odczytano {rows.Count} wierszy z pliku Excel.");
            }
            catch (ArgumentException argEx)
            {
                LogNexo.Blad(argEx);
                Okna.PokazOknoZBledem(uchwyt, $"Bd argumentu odczytu pliku Excel: {argEx.Message}");
                return;
            }
            catch (Exception ex)
            {
                LogNexo.Blad(ex);
                Okna.PokazOknoZBledem(uchwyt, $"Bd odczytu pliku Excel: {ex.Message}");
                return;
            }

            if (rows == null || !rows.Any())
            {
                LogNexo.Informacja("[Import 2.0] Brak danych w podanym pliku Excel.");
                Okna.PokazOknoZBledem(uchwyt, "Wskazany plik nie zawiera danych w wymaganym formacie.");
                return;
            }

            // 2. ?????????? ?????? ? ???'??? ?? ????????? Sfera Manager
            LogNexo.Informacja("[Import 2.0] Rozpoczto wstpn analiz wierszy i walidacj PESEL/Paszport...");
            PreAnalyzeRows(uchwyt, rows);

            // 3. ???????????? ????????? ????? ??????? ??????? nexo (IOknoOperacjiZbiorczej)
            var okno = uchwyt.PodajObiektTypu<IOknoOperacjiZbiorczej>();
            okno.Tytul = "Import 2.0 z HrAppka (Sfera Native)";
            okno.TekstPrzyciskuWykonaj = "Uruchom Import 2.0";

            LogNexo.Informacja("[Import 2.0] Otwieranie nexowego okna operacji zbiorczej.");
            okno.Pokaz(rows, item => ZaimportujRow20(kontekstOperacji, item));

            // 4. ?????????? ???????????? ??? ?????????
            PokazPodsumowanieImportu(kontekstOperacji, rows);
        }

        private bool PodajOSciezkeDoPliku(out string sciezka)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "Pliki danych (*.xlsx;*.csv)|*.xlsx;*.csv|Wszystkie pliki (*.*)|*.*",
                Title = "Import Pracowników - Wybierz plik z danymi (*.xlsx, *.csv)",
                CheckFileExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                sciezka = dialog.FileName;
                return true;
            }
            sciezka = "";
            return false;
        }

        public static List<PracownikImportRow> WczytajExcel(string sciezka)
        {
            return UniversalFileLoader.LoadFile(sciezka);
        }

        public static void SanitizeAddressFields(PracownikImportRow row)
        {
            if (row == null) return;

            string imie = (row.Imie ?? "").Trim();
            string nazwisko = (row.Nazwisko ?? "").Trim();

            // Перевірка Ulica
            if (!string.IsNullOrWhiteSpace(row.Ulica))
            {
                string ulicaTrim = row.Ulica.Trim();
                bool isPib = (!string.IsNullOrEmpty(imie) && ulicaTrim.IndexOf(imie, StringComparison.OrdinalIgnoreCase) >= 0) ||
                             (!string.IsNullOrEmpty(nazwisko) && ulicaTrim.IndexOf(nazwisko, StringComparison.OrdinalIgnoreCase) >= 0);
                bool isPhone = System.Text.RegularExpressions.Regex.IsMatch(ulicaTrim, @"^\+?\d[\d\s\-]{7,}\d$");

                if (isPib || isPhone)
                {
                    row.Ulica = "";
                    row.Ostrzezenia += $"Wykryto błędną treść w polu Ulica ({(isPib ? "Imię/Nazwisko" : "Numer telefonu")}). Pole zostało oczyszczone. ";
                }
            }

            // Перевірка NrDomu
            if (!string.IsNullOrWhiteSpace(row.NrDomu))
            {
                string nrDomuTrim = row.NrDomu.Trim();
                bool isPib = (!string.IsNullOrEmpty(imie) && nrDomuTrim.IndexOf(imie, StringComparison.OrdinalIgnoreCase) >= 0) ||
                             (!string.IsNullOrEmpty(nazwisko) && nrDomuTrim.IndexOf(nazwisko, StringComparison.OrdinalIgnoreCase) >= 0);
                bool isPhone = System.Text.RegularExpressions.Regex.IsMatch(nrDomuTrim, @"^\+?\d[\d\s\-]{7,}\d$");

                if (isPib || isPhone)
                {
                    row.NrDomu = "";
                    row.Ostrzezenia += $"Wykryto błędną treść w polu Nr Domu ({(isPib ? "Imię/Nazwisko" : "Numer telefonu")}). Pole zostało oczyszczone. ";
                }
            }
        }

        private static string GetVal(IDictionary<string, object> dict, params string[] targetKeys)
        {
            if (dict == null) return "";

            // Phase 1: Exact direct key lookup & exact case-insensitive lookup
            foreach (var targetKey in targetKeys)
            {
                if (string.IsNullOrWhiteSpace(targetKey)) continue;

                if (dict.TryGetValue(targetKey, out var directVal) && directVal != null)
                {
                    string s = directVal.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(s)) return s;
                }

                foreach (var kvp in dict)
                {
                    if (kvp.Key == null || kvp.Value == null) continue;
                    if (string.Equals(kvp.Key.Trim(), targetKey.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        string s = kvp.Value.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(s)) return s;
                    }
                }
            }

            // Phase 2: Normalized exact match
            foreach (var targetKey in targetKeys)
            {
                if (string.IsNullOrWhiteSpace(targetKey)) continue;
                string normTarget = RemoveDiacritics(targetKey);

                foreach (var kvp in dict)
                {
                    if (kvp.Key == null || kvp.Value == null) continue;
                    string normKey = RemoveDiacritics(kvp.Key);
                    if (normKey.Equals(normTarget, StringComparison.OrdinalIgnoreCase))
                    {
                        string s = kvp.Value.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(s)) return s;
                    }
                }
            }

            // Phase 3: Substring matching for specific target keys (length >= 5, excluding "pracownik")
            foreach (var targetKey in targetKeys)
            {
                if (string.IsNullOrWhiteSpace(targetKey)) continue;
                string normTarget = RemoveDiacritics(targetKey);
                if (normTarget.Length < 5) continue;
                if (normTarget.Equals("pracownik", StringComparison.OrdinalIgnoreCase)) continue;

                foreach (var kvp in dict)
                {
                    if (kvp.Key == null || kvp.Value == null) continue;
                    string normKey = RemoveDiacritics(kvp.Key);
                    if (normKey.Contains(normTarget))
                    {
                        string s = kvp.Value.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(s)) return s;
                    }
                }
            }
            return "";
        }

        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, InsERT.Moria.ModelDanych.Zawod> _zawodyCache = new System.Collections.Concurrent.ConcurrentDictionary<string, InsERT.Moria.ModelDanych.Zawod>();
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, InsERT.Moria.ModelDanych.FlagaWlasna> _flagiCache = new System.Collections.Concurrent.ConcurrentDictionary<string, InsERT.Moria.ModelDanych.FlagaWlasna>();

        private static InsERT.Moria.ModelDanych.Zawod LookupZawod(IUchwyt uchwyt, string kod)
        {
            if (string.IsNullOrWhiteSpace(kod)) return null;
            return _zawodyCache.GetOrAdd(kod, k =>
                uchwyt.PodajObiektTypu<InsERT.Moria.Kadry.Duze.IZawody>().Dane.Wszystkie().FirstOrDefault(z => z.Kod == k));
        }

        private static InsERT.Moria.ModelDanych.FlagaWlasna LookupFlagaWlasna(IUchwyt uchwyt, string nazwa)
        {
            if (string.IsNullOrWhiteSpace(nazwa)) return null;
            return _flagiCache.GetOrAdd(nazwa, n =>
                uchwyt.PodajObiektTypu<IFlagiWlasne>().Dane.Wszystkie().FirstOrDefault(f => f.Nazwa == n));
        }

        private static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";
            string s = NormalizeForeignLetters(text).Replace(" ", "").Replace("-", "").Replace("/", "").ToLowerInvariant();
            s = s.Replace("ą", "a").Replace("ć", "c").Replace("ę", "e")
                 .Replace("ł", "l").Replace("ń", "n").Replace("ó", "o")
                 .Replace("ś", "s").Replace("ź", "z").Replace("ż", "z");
            return s;
        }

        public static string NormalizeForeignLetters(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            string text = input.Trim();
            text = text.Replace("ñ", "n").Replace("Ñ", "N")
                       .Replace("á", "a").Replace("Á", "A")
                       .Replace("é", "e").Replace("É", "E")
                       .Replace("í", "i").Replace("Í", "I")
                       .Replace("ó", "o").Replace("Ó", "O")
                       .Replace("ú", "u").Replace("Ú", "U")
                       .Replace("ü", "u").Replace("Ü", "U")
                       .Replace("ä", "a").Replace("Ä", "A")
                       .Replace("ö", "o").Replace("Ö", "O")
                       .Replace("ß", "ss")
                       .Replace("à", "a").Replace("è", "e").Replace("ì", "i").Replace("ò", "o").Replace("ù", "u")
                       .Replace("â", "a").Replace("ê", "e").Replace("î", "i").Replace("ô", "o").Replace("û", "u")
                       .Replace("ç", "c").Replace("Ç", "C");
            return text;
        }

        public static void PreAnalyzeRows(IUchwyt uchwyt, List<PracownikImportRow> rows)
        {
            var podmioty = uchwyt.Podmioty();
            var existingPesels = new HashSet<string>(
                podmioty.Dane.WszyscyPracownicy()
                    .Where(p => p.Osoba != null && !string.IsNullOrEmpty(p.Osoba.PESEL))
                    .Select(p => p.Osoba.PESEL),
                StringComparer.OrdinalIgnoreCase
            );

            var existingPassports = new HashSet<string>(
                podmioty.Dane.WszyscyPracownicy()
                    .Where(p => p.Osoba != null && !string.IsNullOrEmpty(p.Osoba.SeriaINumerDokumentuTozsamosci))
                    .Select(p => p.Osoba.SeriaINumerDokumentuTozsamosci),
                StringComparer.OrdinalIgnoreCase
            );

            foreach (var row in rows)
            {
                if (string.IsNullOrWhiteSpace(row.Pesel))
                {
                    if (string.IsNullOrWhiteSpace(row.Paszport))
                    {
                        row.Ostrzezenia += "Brak numeru PESEL oraz Paszportu. ";
                        row.Status = "Brak PESEL i Paszportu";
                    }
                    else
                    {
                        if (existingPassports.Contains(row.Paszport))
                        {
                            row.Ostrzezenia += "Pracownik z tym paszportem już istnieje w bazie. ";
                            row.Status = "Istnieje w bazie";
                            row.DecyzjaDlaIstniejacego = DecyzjaIstniejacyPracownik.AktualizujKartotekeIDodajUmowe;
                        }
                        else
                        {
                            row.Status = "Gotowy (Cudzoziemiec)";
                        }
                    }
                }
                else
                {
                    bool peselOk = PeselHelper.ValidatePesel(row.Pesel, out DateTime? birthDate, out string gender);
                    if (!peselOk)
                    {
                        row.Ostrzezenia += "Niepoprawny numer PESEL. ";
                        row.Status = "Błędny PESEL";
                    }
                    else
                    {
                        row.ResolvedBirthDate = birthDate.Value;
                        row.ResolvedGender = gender;

                        if (!string.IsNullOrWhiteSpace(row.BirthDateExcel) && DateTime.TryParse(row.BirthDateExcel, out DateTime dtExcel))
                        {
                            if (dtExcel.Date != birthDate.Value.Date)
                            {
                                row.Ostrzezenia += $"Data ur. z PESEL ({birthDate.Value:yyyy-MM-dd}). ";
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(row.PlecExcel))
                        {
                            if (!string.Equals(row.PlecExcel, gender, StringComparison.OrdinalIgnoreCase))
                            {
                                row.Ostrzezenia += $"Płeć z PESEL ({gender}). ";
                            }
                        }
                    }

                    if (existingPesels.Contains(row.Pesel))
                    {
                        row.Ostrzezenia += "Pracownik z tym PESEL już istnieje w bazie. ";
                        row.Status = "Istnieje w bazie";
                        row.DecyzjaDlaIstniejacego = DecyzjaIstniejacyPracownik.AktualizujKartotekeIDodajUmowe;
                    }
                    else
                    {
                        row.Status = "Gotowy";
                    }
                }

                if (!string.IsNullOrWhiteSpace(row.BankStr))
                {
                    if (!ParseBankString(row.BankStr, out string iban, out _, out _) || string.IsNullOrWhiteSpace(iban) || iban.Length < 10 || iban.Length > 34)
                    {
                        row.Ostrzezenia += "Niepoprawny numer rachunku bankowego (IBAN). ";
                    }
                }
            }
        }

        private void SprawdzIstniejacyPracownikUmowy(IUchwyt uchwyt, Podmiot? existingOsoba, PracownikImportRow row, string identType, string identVal)
        {
            var decyzja = WpfWarningResolverDialog.RozwiazKonfliktPracownika(row.Nazwisko, row.Imie, identType, identVal);
            row.DecyzjaDlaIstniejacego = decyzja;

            if (decyzja == DecyzjaIstniejacyPracownik.PominWiersz)
            {
                row.Status = "Pominity (Istnieje)";
                row.Ostrzezenia += "Wiersz zosta pominity na danie uytkownika. ";
            }
            else if (decyzja == DecyzjaIstniejacyPracownik.NadpiszUmowe)
            {
                row.Status = "Aktualizacja (Nadpisanie)";
                row.Ostrzezenia += "Uytkownik zdecydowa o nadpisaniu umowy. ";
            }
            else if (decyzja == DecyzjaIstniejacyPracownik.TylkoDodajUmowe)
            {
                row.Status = "Aktualizacja (Nowa umowa, bez zmian)";
                row.Ostrzezenia += "Tylko dodanie nowej umowy. ";
            }
            else
            {
                row.Status = "Aktualizacja (Dane i Nowa umowa)";
                row.Ostrzezenia += "Aktualizacja danych w kartotece oraz wstawienie umowy. ";
            }
        }

        public static WynikOperacji ZaimportujRow20(IKontekstOperacji? kontekstOperacji, PracownikImportRow item, IUchwyt? fallbackUchwyt = null)
        {
            var uchwyt = kontekstOperacji?.Uchwyt ?? fallbackUchwyt;
            if (uchwyt == null) throw new ArgumentNullException(nameof(fallbackUchwyt), "Uchwyt Sfera SDK jest wymagany.");
            LogNexo.Informacja($"[Import 2.0] Rozpoczto przetwarzanie wiersza: {item.Imie} {item.Nazwisko} (PESEL: {item.Pesel}, Paszport: {item.Paszport})");

            if (item.Status.StartsWith("Pominity"))
            {
                LogNexo.Informacja($"[Import 2.0] Pominito pracownika {item.Imie} {item.Nazwisko} - pracownik ju istnieje.");
                return WynikOperacji.ZakonczonaPowodzeniem("Pominito - pracownik ju istnieje.");
            }

            try
            {
                DateTime? finalBirthDate = null;
                string finalGender = "";
                bool hasPesel = PeselHelper.ValidatePesel(item.Pesel, out DateTime? birthDate, out string gender);
                if (hasPesel)
                {
                    finalBirthDate = birthDate;
                    finalGender = gender;
                }

                if (finalBirthDate == null && !string.IsNullOrWhiteSpace(item.BirthDateExcel))
                {
                    if (DateTime.TryParse(item.BirthDateExcel, out DateTime parsedDt))
                        finalBirthDate = parsedDt;
                }
                if (string.IsNullOrEmpty(finalGender))
                {
                    finalGender = item.PlecExcel;
                }

                var podmiotyMgr = uchwyt.Podmioty();

                // A. ????? ?? 3-??????? ????????????? ?????? Osoba / Pracownik
                Podmiot? existingOsoba = null;
                if (!string.IsNullOrWhiteSpace(item.Pesel))
                {
                    existingOsoba = podmiotyMgr.Dane.WszystkieOsoby()
                        .FirstOrDefault(p => p.Osoba != null && p.Osoba.PESEL == item.Pesel);
                }
                else if (!string.IsNullOrWhiteSpace(item.Paszport))
                {
                    existingOsoba = podmiotyMgr.Dane.WszystkieOsoby()
                        .FirstOrDefault(p => p.Osoba != null && p.Osoba.SeriaINumerDokumentuTozsamosci == item.Paszport);
                }

                IPodmiot pracownikBO;
                if (existingOsoba != null)
                {
                    if (existingOsoba.Osoba.Pracownik != null)
                    {
                        LogNexo.Informacja($"[Import 2.0] Ga 1: Znaleziono istniejcego pracownika (ID: {existingOsoba.Id}). Znajdz().");
                        pracownikBO = podmiotyMgr.Znajdz(existingOsoba);
                    }
                    else
                    {
                        LogNexo.Informacja($"[Import 2.0] Ga 2: Znaleziono osob (ID: {existingOsoba.Id}) bez roli pracownika. PrzeksztalcOsobeNaPracownika().");
                        pracownikBO = podmiotyMgr.PrzeksztalcOsobeNaPracownika(existingOsoba);
                    }
                }
                else
                {
                    LogNexo.Informacja("[Import 2.0] Ga 3: Tworzenie nowego pracownika via UtworzPracownika() z AutoSymbol().");
                    pracownikBO = podmiotyMgr.UtworzPracownika();
                    pracownikBO.AutoSymbol();
                }

                Podmiot savedEmployeePodmiot;

                using (pracownikBO)
                {
                    var podmiot = pracownikBO.Dane;
                    var osoba = podmiot.Osoba;

                    // 1. Organ Podatkowy (Urzd Skarbowy)
                    if (!string.IsNullOrWhiteSpace(item.UrzadSkarbowy))
                    {
                        LogNexo.Informacja($"[Import 2.0] Resolving Urzd Skarbowy '{item.UrzadSkarbowy}'");
                        var organ = ResolveOrganPodatkowy(uchwyt, item.UrzadSkarbowy);
                        if (organ != null)
                        {
                            podmiot.OrganPodatkowy = organ;
                            LogNexo.Informacja($"[Import 2.0] Przypisano Urzd Skarbowy: {organ.Firma?.Nazwa ?? item.UrzadSkarbowy}");
                        }
                        else
                        {
                            podmiot.OrganPodatkowy = null;
                            string note = $"US wskazany pracownikiem - \"{item.UrzadSkarbowy.Trim()}\"";
                            if (string.IsNullOrWhiteSpace(podmiot.Uwagi))
                            {
                                podmiot.Uwagi = note;
                            }
                            else if (!podmiot.Uwagi.Contains("US wskazany pracownikiem"))
                            {
                                podmiot.Uwagi = (podmiot.Uwagi + "; " + note).Trim();
                                if (podmiot.Uwagi.Length > 256)
                                    podmiot.Uwagi = podmiot.Uwagi.Substring(0, 256);
                            }
                            item.Ostrzezenia += $"Nie znaleziono urzdu skarbowego '{item.UrzadSkarbowy}' w bazie. Pozostawiono pustym, zapisano w uwagach. ";
                            LogNexo.Informacja($"[Import 2.0] Nie znaleziono urzdu skarbowego '{item.UrzadSkarbowy}'. Zapisano w uwagach.");
                        }
                    }
                    else
                    {
                        podmiot.OrganPodatkowy = null;
                    }

                    // 2. Personal Details & Name Splitting (SplitNames with Foreign Letter Normalization)
                    string normImie = NormalizeForeignLetters(PeselHelper.NormalizeWhitespace(item.Imie));
                    string normDrugieImie = NormalizeForeignLetters(PeselHelper.NormalizeWhitespace(item.DrugieImie));
                    string normNazwisko = NormalizeForeignLetters(PeselHelper.NormalizeWhitespace(item.Nazwisko));

                    string finalImie = "";
                    string finalDrugieImie = "";

                    if (!string.IsNullOrEmpty(normImie))
                    {
                        var imieTokens = normImie.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        finalImie = imieTokens[0].ToUpper();

                        string extractedSecond = imieTokens.Length > 1 
                            ? string.Join(" ", imieTokens.Skip(1)).ToUpper() 
                            : "";
                        string excelSecond = normDrugieImie.ToUpper();

                        if (string.IsNullOrEmpty(extractedSecond) && string.IsNullOrEmpty(excelSecond))
                        {
                            finalDrugieImie = "";
                        }
                        else if (string.IsNullOrEmpty(extractedSecond))
                        {
                            finalDrugieImie = excelSecond;
                        }
                        else if (string.IsNullOrEmpty(excelSecond))
                        {
                            finalDrugieImie = extractedSecond;
                        }
                        else
                        {
                            if (string.Equals(extractedSecond, excelSecond, StringComparison.OrdinalIgnoreCase))
                            {
                                finalDrugieImie = extractedSecond;
                            }
                            else
                            {
                                finalDrugieImie = ZapytajOUzytkownikaODrugieImie(normNazwisko.ToUpper(), finalImie, extractedSecond, excelSecond);
                            }
                        }
                    }
                    else
                    {
                        finalDrugieImie = normDrugieImie.ToUpper();
                    }

                    osoba.Imie = finalImie;
                    osoba.DrugieImie = finalDrugieImie;
                    osoba.Nazwisko = normNazwisko.ToUpper();
                    osoba.PESEL = item.Pesel;
                    if (finalBirthDate.HasValue)
                        osoba.DataUrodzenia = finalBirthDate.Value;

                    if (!string.IsNullOrEmpty(finalGender))
                    {
                        osoba.Plec = (byte)GetPlec(finalGender);
                    }

                    LogNexo.Informacja($"[Import 2.0] Ustawiono dane osobowe: Imi='{osoba.Imie}', DrugieImi='{osoba.DrugieImie}', Nazwisko='{osoba.Nazwisko}'");

                    // 3. Citizenship & Country
                    Panstwo? resolvedPanstwo = null;
                    if (!string.IsNullOrWhiteSpace(item.Obywatelstwo))
                    {
                        LogNexo.Informacja($"[Import 2.0] Resolving Obywatelstwo dla: '{item.Obywatelstwo}'");
                        var (panstwo, obywatelstwo, _) = ZusCitizenshipManager.Instance.ResolveSfera(uchwyt, item.Obywatelstwo);
                        resolvedPanstwo = panstwo;
                        if (obywatelstwo != null)
                        {
                            osoba.Pracownik.Obywatelstwo = obywatelstwo;
                            LogNexo.Informacja($"[Import 2.0] Przypisano obywatelstwo: {obywatelstwo.Nazwa}");
                        }
                    }

                    // 4. Address Details (ZUS Płatnik & Nexo DB Schema Compliant)
                    if (!string.IsNullOrWhiteSpace(item.Miejscowosc) || !string.IsNullOrWhiteSpace(item.Ulica) || !string.IsNullOrWhiteSpace(item.NrDomu))
                    {
                        var adres = podmiot.AdresPodstawowy ?? pracownikBO.DodajAdres(uchwyt.TypyAdresu().DaneDomyslne.Glowny);
                        adres.Nazwa = "Adres zamieszkania";

                        string rawHouseApt = string.IsNullOrWhiteSpace(item.NrLokalu)
                            ? item.NrDomu
                            : $"{item.NrDomu}/{item.NrLokalu}";

                        var norm = AddressNormalizer.Normalize(
                            rawStreet: item.Ulica,
                            rawHouseAndApt: rawHouseApt,
                            rawCity: item.Miejscowosc,
                            rawPostalCode: item.KodPocztowy,
                            rawGmina: item.Gmina,
                            rawPowiat: item.Powiat,
                            rawWojewodztwo: item.Wojewodztwo,
                            rawCountry: item.Kraj
                        );

                        adres.Szczegoly.Ulica = norm.Ulica;
                        adres.Szczegoly.NrDomu = norm.NrDomu;
                        adres.Szczegoly.NrLokalu = norm.NrLokalu;
                        adres.Szczegoly.Miejscowosc = norm.Miejscowosc;
                        adres.Szczegoly.KodPocztowy = norm.KodPocztowy;
                        adres.Szczegoly.Poczta = norm.Poczta;
                        adres.Panstwo = ZnajdzPanstwo(uchwyt, norm.Kraj);

                        var woj = ZnajdzWojewodztwo(uchwyt, item.Wojewodztwo);
                        if (woj != null) adres.Szczegoly.Wojewodztwo = woj;

                        var pow = ZnajdzPowiat(uchwyt, woj, item.Powiat);
                        if (pow != null)
                        {
                            var gmina = ZnajdzGmine(uchwyt, pow, norm.Gmina);
                            if (gmina != null) adres.Szczegoly.Gmina = gmina;
                        }
                        else
                        {
                            var gmina = ZnajdzGmine(uchwyt, null, norm.Gmina);
                            if (gmina != null) adres.Szczegoly.Gmina = gmina;
                        }

                        LogNexo.Informacja($"[Import 2.0] Znormalizowano adres: Ulica='{norm.Ulica}', Dom='{norm.NrDomu}', Lok='{norm.NrLokalu}', Miasto='{norm.Miejscowosc}', Kod='{norm.KodPocztowy}'");
                    }

                    // 5. Contacts (Email & Telefon z dedupilakcj)
                    if (!string.IsNullOrWhiteSpace(item.Email))
                    {
                        string trimmedEmail = item.Email.Trim();
                        var existingEmail = podmiot.Kontakty.FirstOrDefault(k => k.Rodzaj.Id == uchwyt.RodzajeKontaktu().DaneDomyslne.Email.Id);
                        if (existingEmail != null)
                        {
                            existingEmail.Wartosc = trimmedEmail;
                            existingEmail.Podstawowy = true;
                        }
                        else
                        {
                            var k = new Kontakt();
                            podmiot.Kontakty.Add(k);
                            k.UniqueId = Guid.NewGuid();
                            k.Rodzaj = uchwyt.RodzajeKontaktu().DaneDomyslne.Email;
                            k.Wartosc = trimmedEmail;
                            k.Podstawowy = true;
                        }
                    }

                    string cleanedPhone = "";
                    if (!string.IsNullOrWhiteSpace(item.Telefon))
                    {
                        var sb = new StringBuilder();
                        for (int i = 0; i < item.Telefon.Length; i++)
                        {
                            char c = item.Telefon[i];
                            if (char.IsDigit(c)) sb.Append(c);
                            else if (c == '+' && sb.Length == 0) sb.Append(c);
                        }
                        cleanedPhone = sb.ToString();
                    }

                    if (!string.IsNullOrWhiteSpace(cleanedPhone))
                    {
                        var existingPhone = podmiot.Kontakty.FirstOrDefault(k => k.Rodzaj.Id == uchwyt.RodzajeKontaktu().DaneDomyslne.Telefon.Id);
                        if (existingPhone != null)
                        {
                            existingPhone.Wartosc = cleanedPhone;
                            existingPhone.Podstawowy = true;
                        }
                        else
                        {
                            var k = new Kontakt();
                            podmiot.Kontakty.Add(k);
                            k.UniqueId = Guid.NewGuid();
                            k.Rodzaj = uchwyt.RodzajeKontaktu().DaneDomyslne.Telefon;
                            k.Wartosc = cleanedPhone;
                            k.Podstawowy = true;
                        }
                    }

                    // 6. Identyfikacja Dokumentu Tożsamości (zgodna z priorytetami ZUS i krajem wydania)
                    var docResult = DocumentIdentificationService.DetermineDocument(item, resolvedPanstwo);
                    DocumentIdentificationService.ApplyToEmployee(podmiot, docResult);
                    if (docResult.HasDocument)
                    {
                        LogNexo.Informacja($"[Import 2.0] Przypisano dokument tożsamości: {docResult.DocumentType} {docResult.DocumentNumber}, Kraj: {docResult.CountryOfIssue?.Nazwa}");
                    }

                    // 7. Gratyfikant Specyfics & NFZ
                    var pracownikGr = osoba.Pracownik.PracownikGr;
                    DateTime contractStart = DateTime.Today;
                    if (!string.IsNullOrWhiteSpace(item.DataRozpoczecia) && DateTime.TryParse(item.DataRozpoczecia, out DateTime parsedStart))
                    {
                        contractStart = parsedStart;
                    }

                    pracownikGr.DaneWdrozenioweGr.MiesiacPierwszejWyplaty = new DateTime(contractStart.Year, contractStart.Month, 1);
                    pracownikGr.DataPrzystapieniaDoNFZ = contractStart;
                    pracownikGr.Wyksztalcenie = (byte)0;

                    if (!string.IsNullOrWhiteSpace(item.NfzStr))
                    {
                        var nfz = MappedNFZ(item.NfzStr);
                        if (nfz.HasValue)
                        {
                            pracownikGr.OddzialNFZ = (byte)nfz.Value;
                        }
                    }

                    // 8. Bank Account Creation & Primary Assignment
                    if (!string.IsNullOrWhiteSpace(item.BankStr))
                    {
                        if (ParseBankString(item.BankStr, out string iban, out string currencySymbol, out string bankName))
                        {
                            var currency = uchwyt.PodajObiektTypu<IWaluty>().Dane.Wszystkie().FirstOrDefault(w => w.Symbol == currencySymbol)
                                           ?? uchwyt.PodajObiektTypu<IWaluty>().Dane.Wszystkie().FirstOrDefault(w => w.Symbol == "PLN");

                            string cleanIban = iban.Replace(" ", "").Replace("-", "");
                            var existingRachunek = podmiot.Rachunki.FirstOrDefault(r => r.Numer != null && r.Numer.Replace(" ", "").Replace("-", "") == cleanIban);

                            if (existingRachunek != null)
                            {
                                existingRachunek.Nazwa = bankName;
                                existingRachunek.Waluta = currency;
                                podmiot.RachunekPodstawowy = existingRachunek;
                            }
                            else
                            {
                                var rachunek = new RachunekBankowy();
                                podmiot.Rachunki.Add(rachunek);
                                rachunek.UniqueId = Guid.NewGuid();
                                rachunek.Numer = iban;
                                rachunek.Nazwa = bankName;
                                rachunek.Waluta = currency;
                                podmiot.RachunekPodstawowy = rachunek;
                            }
                            LogNexo.Informacja($"[Import 2.0] Utworzono/przypisano rachunek bankowy: {iban} ({bankName})");
                        }
                    }

                    // B. Zapisanie Kartki Pracownika z głęboką diagnostyką UoW
                    if (!pracownikBO.Zapisz())
                    {
                        string errors = GetDetailedUowErrors(pracownikBO, uchwyt);
                        if (errors.Contains("rachunek bankowy", StringComparison.OrdinalIgnoreCase))
                        {
                            LogNexo.Informacja($"[Import 2.0] IBAN jest przypisany do innego podmiotu w Nexo. Czyszczenie rachunku dla {item.Imie} {item.Nazwisko}...");
                            podmiot.Rachunki.Clear();
                            podmiot.RachunekPodstawowy = null;
                            item.Ostrzezenia += "Rachunek bankowy należy do innego podmiotu w базе. ";

                            if (!pracownikBO.Zapisz())
                            {
                                errors = GetDetailedUowErrors(pracownikBO, uchwyt);
                                LogNexo.Informacja($"[Import 2.0] Błąd zapisu karty pracownika {item.Imie} {item.Nazwisko}: {errors}");
                                item.Status = "Błąd zapisu pracownika";
                                item.Ostrzezenia += errors + " ";
                                return WynikOperacji.ZakonczonaNiepowodzeniem($"Błąd zapisu karty pracownika: {errors}");
                            }
                        }
                        else
                        {
                            LogNexo.Informacja($"[Import 2.0] Błąd zapisu karty pracownika {item.Imie} {item.Nazwisko}: {errors}");
                            item.Status = "Błąd zapisu pracownika";
                            item.Ostrzezenia += errors + " ";
                            return WynikOperacji.ZakonczonaNiepowodzeniem($"Błąd zapisu karty pracownika: {errors}");
                        }
                    }

                    savedEmployeePodmiot = pracownikBO.Dane;
                    LogNexo.Informacja($"[Import 2.0] Pomylnie zapisano kart pracownika: {item.Imie} {item.Nazwisko} (ID={savedEmployeePodmiot.Id})");
                }

                // C. ????????? ?? ???????????? umowy cywilnoprawnej (Umowa Cywilnoprawna)
                if (item.Status == "Aktualizacja (Dane)")
                {
                    LogNexo.Informacja($"[Import 2.0] Pominito tworzenie umowy dla {item.Imie} {item.Nazwisko} - umowa ju istnieje.");
                }
                else if (!string.IsNullOrWhiteSpace(item.DataRozpoczecia) && DateTime.TryParse(item.DataRozpoczecia, out DateTime dtStart))
                {
                    DateTime? dtEnd = DateTime.TryParse(item.DataZakonczenia, out DateTime endParsed) ? endParsed : null;
                    DateTime dtSign = DateTime.TryParse(item.DataZawarcia, out DateTime signParsed) ? signParsed : dtStart;

                    var umowaZlecenieTyp = uchwyt.TypyUmowPracowniczych().Dane.Wszystkie().FirstOrDefault(x => x.Nazwa == "Umowa zlecenie");
                    if (umowaZlecenieTyp == null)
                    {
                        string noTypeMsg = "Nie znaleziono typu umowy 'Umowa zlecenie' w bazie nexo.";
                        LogNexo.Informacja($"[Import 2.0] {noTypeMsg}");
                        item.Status = "Bd (Brak typu umowy)";
                        item.Ostrzezenia += noTypeMsg + " ";
                        return WynikOperacji.ZakonczonaNiepowodzeniem(noTypeMsg);
                    }

                    using (var umowaBO = uchwyt.UmowyPracowniczeGr().Utworz(umowaZlecenieTyp))
                    {
                        var u = umowaBO.Dane;
                        u.Pracownik = savedEmployeePodmiot.Osoba.Pracownik;

                        u.RodzajUmowyPracowniczej = umowaZlecenieTyp.DozwoloneRodzajeUmowyPracowniczej
                            .FirstOrDefault(x => string.Equals(x.Nazwa, "Na czas określony", StringComparison.OrdinalIgnoreCase)) 
                            ?? umowaZlecenieTyp.DozwoloneRodzajeUmowyPracowniczej.FirstOrDefault();
                        u.WzorzecTresciUmowyPracowniczej = uchwyt.TresciWydrukuRTF().DaneDomyslne.UCPZlecenieStale;

                        if (u.OkresObowiazywania == null)
                            u.OkresObowiazywania = new OkresWymaganyLewostronnie();
                        u.OkresObowiazywania.DataPoczatkowa = dtStart;
                        u.OkresObowiazywania.DataKoncowa = dtEnd;
                        u.DataSporzadzenia = dtSign;

                        // Calculate age & student status (ZUS / PIT)
                        int age = PeselHelper.CalculateAge(finalBirthDate ?? DateTime.MinValue, dtStart);
                        bool studentFlag = string.Equals(item.JestStudentem, "Tak", StringComparison.OrdinalIgnoreCase);
                        bool ageUnder26 = age < 26;
                        bool isStudentUnder26 = studentFlag && ageUnder26;
                        bool isZusSubject = !isStudentUnder26;

                        u.RaportZUS = (byte)(isStudentUnder26 ? InsERT.Moria.Kadry.TypRaportuZUS.Brak : InsERT.Moria.Kadry.TypRaportuZUS.RCA);
                        u.NaliczaSkladkiEmerytalne = isZusSubject;
                        u.NaliczaSkladkiRentowe = isZusSubject;
                        u.NaliczaSkladkiZdrowotne = isZusSubject;
                        u.NaliczaSkladkiWypadkowe = isZusSubject;
                        u.NaliczaSkladkiFP = isZusSubject;
                        u.NaliczaSkladkiFGSP = isZusSubject;
                        u.NaliczaSkladkiChorobowe = false;

                        u.StosujZwolnienieZPITDo26Lat = null;

                        // Limits
                        if (u.NaliczaSkladkiZdrowotne) u.NieOgraniczaZdrowotnego = false;
                        if (u.NaliczaSkladkiChorobowe) u.NieOgraniczaChorobowego = false;
                        if (u.NaliczaSkladkiFP) u.NaliczaSkladkiFPTylkoPowyzejMinimalnej = isZusSubject;

                        // PIT & Relief
                        try
                        {
                            var targetZaliczka = (byte)(ageUnder26 
                                ? SposobPobieraniaZaliczkiNaPodatekNaUmowie.WgUstawienPracownika 
                                : SposobPobieraniaZaliczkiNaPodatekNaUmowie.NiePobierajDoKwotyWolnej);
                            if (u.ZaliczkaNaPodatek != targetZaliczka) u.ZaliczkaNaPodatek = targetZaliczka;

                            var targetUlga = (byte)SposobOdliczaniaUlgiPodatkowejNaUmowie.WgUstawienPracownika;
                            if (u.UlgaPodatkowa != targetUlga) u.UlgaPodatkowa = targetUlga;
                        }
                        catch (Exception ex)
                        {
                            LogNexo.Informacja($"[Import 2.0] Ignorowano bd ustawiania PIT: {ex.Message}");
                        }

                        // Schedule & Wage rates
                        try
                        {
                            var targetWyplata = (byte)WyplacanieWynagrodzenia.CoMiesiac;
                            if (u.WyplacaWynagrodzenie != targetWyplata) u.WyplacaWynagrodzenie = targetWyplata;
                            if (u.DataRachunkuPrzesuniecie != 25) u.DataRachunkuPrzesuniecie = 25;
                            if (u.PrzesuniecieNaliczenia != 0) u.PrzesuniecieNaliczenia = 0;
                        }
                        catch (Exception ex)
                        {
                            LogNexo.Informacja($"[Import 2.0] Ignorowano bd harmonogramu wypat: {ex.Message}");
                        }

                        foreach (var sp in u.SkladnikiPlacowe)
                        {
                            foreach (var val in sp.Wartosci)
                            {
                                val.Wartosc = 0.00m;
                            }
                        }

                        // Dzia / Department
                        var subDept = DepartmentResolver.ResolveDepartment(uchwyt, item.DzialNazwa);
                        if (subDept != null)
                        {
                            var st = new StanowiskoWUmowiePracowniczej();
                            u.Stanowiska.Add(st);
                            st.DataOd = dtStart;
                            st.JednostkaOrganizacyjna = subDept;
                            st.LP = 1;
                        }

                        // Occupation / Kod zawodu
                        if (!string.IsNullOrWhiteSpace(item.KodZawodu))
                        {
                            var zawod = LookupZawod(uchwyt, item.KodZawodu);
                            if (zawod != null) u.Zawod = zawod;
                        }

                        // Flaga wlasna (OPODATKOWANE / NIEOPODATKOWANE)
                        try
                        {
                            string flagName = isStudentUnder26 ? "NIEOPODATKOWANE" : "OPODATKOWANE";
                            var flag = LookupFlagaWlasna(uchwyt, flagName);
                            if (flag != null) u.FlagaWlasna = flag;
                            else item.Ostrzezenia += $"Nie znaleziono flagi własnej '{flagName}' w bazie. ";
                        }
                        catch (Exception ex)
                        {
                            item.Ostrzezenia += $"Błąd przypisywania flagi własnej: {ex.Message}. ";
                        }

                        if (!umowaBO.Zapisz())
                        {
                            string contractErrors = GetDetailedUowErrors(umowaBO, uchwyt);
                            LogNexo.Informacja($"[Import 2.0] Bd zapisu umowy dla {item.Imie} {item.Nazwisko}: {contractErrors}");
                            item.Status = "Bd zapisu umowy";
                            item.Ostrzezenia += contractErrors + " ";
                            return WynikOperacji.ZakonczonaNiepowodzeniem($"Karta pracownika utworzona, lecz bd zapisu umowy: {contractErrors}");
                        }

                        LogNexo.Informacja($"[Import 2.0] Pomylnie zapisano umow dla {item.Imie} {item.Nazwisko}");

                        // ZUS Declaration ZUA (if RCA)
                        if (u.RaportZUS == (byte)InsERT.Moria.Kadry.TypRaportuZUS.RCA)
                        {
                            DodajZgloszenieZus(uchwyt, savedEmployeePodmiot, u, item);
                        }
                    }
                }

                if (!string.IsNullOrEmpty(item.Ostrzezenia))
                {
                    item.Status = "Zaimportowano (Z ostrzeeniami)";
                    LogNexo.Informacja($"[Import 2.0] Zakoczono import {item.Imie} {item.Nazwisko} z ostrzeeniami: {item.Ostrzezenia}");
                    return WynikOperacji.ZakonczonaPowodzeniem($"Zaimportowano z ostrzeeniami: {item.Ostrzezenia}");
                }

                item.Status = "Zaimportowano (Import 2.0)";
                LogNexo.Informacja($"[Import 2.0] Zakoczono pomylnie import dla {item.Imie} {item.Nazwisko}");
                return WynikOperacji.ZakonczonaPowodzeniem("Import 2.0 zakoczony sukcesem.");
            }
            catch (ArgumentException argEx)
            {
                LogNexo.Blad(argEx);
                item.Status = "Bd argumentu";
                item.Ostrzezenia += $"Wyjtek argumentu: {argEx.Message} ";
                return WynikOperacji.ZakonczonaNiepowodzeniem($"Wyjtek argumentu: {argEx.Message}");
            }
            catch (Exception ex)
            {
                LogNexo.Blad(ex);
                item.Status = "Wyjtek";
                item.Ostrzezenia += $"Wyjtek: {ex.Message} ";
                return WynikOperacji.ZakonczonaNiepowodzeniem($"Wyjtek podczas importu 2.0: {ex.Message}");
            }
        }

        private static void DodajZgloszenieZus(IUchwyt uchwyt, Podmiot employeePodmiot, UmowaPracownicza baseUmowa, PracownikImportRow item)
        {
            try
            {
                var dekZusMgr = uchwyt.PodajObiektTypu<IDeklaracjeZUS>();
                var existingDecl = dekZusMgr.Dane.Wszystkie()
                    .Where(d => d.RodzajDeklaracji == (byte)RodzajDeklaracjiZUS.Zgloszeniowa 
                             && d.StatusWysylki == (byte)StatusWysylkiDeklaracjiZUS.Brak
                             && !d.WyslanyDoPlatnika)
                    .OrderByDescending(d => d.Id)
                    .FirstOrDefault();

                IDeklaracjaZUS deklBO = existingDecl != null ? dekZusMgr.Znajdz(existingDecl) : dekZusMgr.UtworzDeklaracjeZgloszeniowa();

                using (deklBO)
                {
                    bool alreadyAdded = deklBO.Dane.DeklaracjeZgloszeniowe
                        .Any(z => z.UmowaPracownicza != null && z.UmowaPracownicza.Id == baseUmowa.Id);

                    if (!alreadyAdded)
                    {
                        var zgloszenie = new DeklaracjaZgloszeniowa();
                        deklBO.Dane.DeklaracjeZgloszeniowe.Add(zgloszenie);

                        var podmioty = uchwyt.Podmioty();
                        var umowy = uchwyt.UmowyPracowniczeGr();
                        var freshEmployeePodmiot = podmioty.Dane.WszystkieOsoby().FirstOrDefault(p => p.Id == employeePodmiot.Id);
                        var freshContract = umowy.Dane.Wszystkie().FirstOrDefault(u => u.Id == baseUmowa.Id);
                        var freshEmployee = freshEmployeePodmiot?.Osoba?.Pracownik;

                        if (freshContract != null && freshEmployee != null)
                        {
                            if (freshEmployee.Osoba == null && freshEmployeePodmiot?.Osoba != null)
                            {
                                freshEmployee.Osoba = freshEmployeePodmiot.Osoba;
                            }

                            zgloszenie.UmowaPracownicza = freshContract;
                            zgloszenie.Pracownik = freshEmployee;
                            zgloszenie.Typ = (byte)TypDeklaracjiZgloszeniowej.ZUA;
                            zgloszenie.ImieNazwisko = $"{freshEmployee.Osoba.Imie} {freshEmployee.Osoba.Nazwisko}".ToUpper().Trim();

                            Type? daneZuaType = null;
                            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                            {
                                daneZuaType = asm.GetType("InsERT.Moria.DeklaracjeZUS.DaneZUA");
                                if (daneZuaType != null) break;
                            }

                            if (daneZuaType == null)
                            {
                                try
                                {
                                    var privateAsm = System.Reflection.Assembly.LoadFrom(@"C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo\KLINEX SP_Z_O_Oa4bfb87390e94eea9\Binaries\InsERT.Moria.API.Private.dll");
                                    daneZuaType = privateAsm.GetType("InsERT.Moria.DeklaracjeZUS.DaneZUA");
                                }
                                catch (Exception ex)
                                {
                                    LogNexo.Informacja($"[Import 2.0] Bd adowania API.Private.dll dla ZUS: {ex.Message}");
                                }
                            }

                            if (daneZuaType != null)
                            {
                                var daneZUA = Activator.CreateInstance(daneZuaType);
                                daneZuaType.GetProperty("TYPZgloszenia")?.SetValue(daneZUA, 0);
                                DateTime obligationDate = freshContract.OkresObowiazywania?.DataPoczatkowa ?? DateTime.Today;
                                daneZuaType.GetProperty("DataObowiazkuUbezpieczenSpoleczych")?.SetValue(daneZUA, obligationDate);
                                daneZuaType.GetProperty("DataObowiazkuUbezpieczenZdrowotnych")?.SetValue(daneZUA, obligationDate);

                                daneZuaType.GetProperty("PlacSkladkaEmerytalna")?.SetValue(daneZUA, freshContract.NaliczaSkladkiEmerytalne);
                                daneZuaType.GetProperty("PlacSkladkaRentowa")?.SetValue(daneZUA, freshContract.NaliczaSkladkiRentowe);
                                daneZuaType.GetProperty("PlacSkladkaChorobowa")?.SetValue(daneZUA, freshContract.NaliczaSkladkiChorobowe);
                                daneZuaType.GetProperty("PlacSkladkaWypadkowa")?.SetValue(daneZUA, freshContract.NaliczaSkladkiWypadkowe);

                                var wypelnijMethod = deklBO.GetType().GetMethod("WypelnijDeklaracjeZgloszeniona",
                                    System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                                if (wypelnijMethod != null)
                                {
                                    wypelnijMethod.Invoke(deklBO, new object[] { zgloszenie, daneZUA });
                                }
                                else
                                {
                                    item.Ostrzezenia += "Nie odnaleziono metody WypelnijDeklaracjeZgloszeniona. ";
                                }
                            }

                            if (!deklBO.Zapisz())
                            {
                                var bledy = uchwyt.PodajBledy(deklBO);
                                var msg = string.Join("; ", bledy.Select(x => $"[{x.Waznosc}] {x.Tresc}"));
                                item.Ostrzezenia += $"Nie udao si zapisa zgoszenia ZUS: {msg}. ";
                                LogNexo.Informacja($"[Import 2.0] Nie udao si zapisa zgoszenia ZUS: {msg}");
                            }
                            else
                            {
                                LogNexo.Informacja($"[Import 2.0] Dodano umow pracownika {item.Imie} {item.Nazwisko} do zgoszenia ZUS (ZUA).");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                item.Ostrzezenia += $"Bd podczas tworzenia zgoszenia ZUS ZUA: {ex.Message}. ";
                LogNexo.Blad(ex);
            }
        }

        private static string GetDetailedUowErrors(InsERT.Mox.ObiektyBiznesowe.IObiektBiznesowy bo, IUchwyt uchwyt)
        {
            var bledy = uchwyt.PodajBledy(bo);
            var msgList = bledy != null 
                ? bledy.Select(x => $"[{x.Waznosc}] {x.Tresc}").ToList() 
                : new List<string>();

            try
            {
                var uow = ((InsERT.Mox.BusinessObjects.IGetUnitOfWork)bo).UnitOfWork;
                foreach (var participant in uow.Participants.OfType<InsERT.Mox.BusinessObjects.IBusinessObject>())
                {
                    var dd = (InsERT.Mox.DataAccess.IGetDataDomain)participant;
                    dd.DataDomain.TraverseData(participant.Data, (obj) =>
                    {
                        if (obj is InsERT.Mox.Validation.ITypedDataErrorInfo errInfo)
                        {
                            foreach (var err in errInfo.Errors)
                                msgList.Add($"[{obj.GetType().Name}] {err}");
                            foreach (var memberErr in errInfo.MemberErrors)
                                msgList.Add($"[{obj.GetType().Name}.{memberErr.Key}] {string.Join(", ", memberErr)}");
                        }
                        return true;
                    });
                }
            }
            catch (Exception ex)
            {
                msgList.Add($"[DiagError] {ex.Message}");
            }

            return msgList.Any(m => !string.IsNullOrWhiteSpace(m)) 
                ? string.Join("; ", msgList.Where(m => !string.IsNullOrWhiteSpace(m)).Distinct()) 
                : "Bd zapisu w Sfera (walidacja nexo).";
        }

        private static Plec GetPlec(string plecStr)
        {
            if (string.IsNullOrWhiteSpace(plecStr)) return Plec.Nieokreslona;
            string lower = plecStr.ToLower();
            if (lower.Contains("kob") || lower.Contains("fem")) return Plec.Kobieta;
            if (lower.Contains("m") || lower.Contains("mez") || lower.Contains("mal")) return Plec.Mezczyzna;
            return Plec.Nieokreslona;
        }

        private static Panstwo ZnajdzPanstwo(IUchwyt uchwyt, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) name = "Polska";
            string norm = name.ToUpper();
            var p = uchwyt.Panstwa().Dane.Wszystkie().FirstOrDefault(x => x.Nazwa.ToUpper() == norm || x.KodPanstwaUE.ToUpper() == norm);
            if (p == null)
            {
                p = uchwyt.Panstwa().Dane.Wszystkie().FirstOrDefault(x => x.Nazwa.ToUpper() == "POLSKA");
            }
            return p!;
        }

        private static Wojewodztwo? ZnajdzWojewodztwo(IUchwyt uchwyt, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            string norm = name.ToUpper();
            return uchwyt.Wojewodztwa().Dane.Wszystkie().FirstOrDefault(w => w.Nazwa.ToUpper() == norm);
        }

        private static Powiat? ZnajdzPowiat(IUchwyt uchwyt, Wojewodztwo? woj, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            string norm = name.ToUpper();
            if (woj != null)
            {
                return uchwyt.Powiaty().Dane.Wszystkie().FirstOrDefault(p => p.Wojewodztwo.Id == woj.Id && p.Nazwa.ToUpper() == norm);
            }
            return uchwyt.Powiaty().Dane.Wszystkie().FirstOrDefault(p => p.Nazwa.ToUpper() == norm);
        }

        private static Gmina? ZnajdzGmine(IUchwyt uchwyt, Powiat? pow, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            string norm = name.ToUpper();
            if (pow != null)
            {
                return uchwyt.Gminy().Dane.Wszystkie().FirstOrDefault(g => g.Powiat.Id == pow.Id && g.Nazwa.ToUpper() == norm);
            }
            return uchwyt.Gminy().Dane.Wszystkie().FirstOrDefault(g => g.Nazwa.ToUpper() == norm);
        }

        private static OddzialNFZ? MappedNFZ(string nfzName)
        {
            if (string.IsNullOrWhiteSpace(nfzName)) return null;
            string lower = nfzName.ToLower();
            if (lower.Contains("lubel") || lower.Contains("lublin")) return OddzialNFZ.lublin;
            if (lower.Contains("lubus") || lower.Contains("zielon")) return OddzialNFZ.zielona;
            if (lower.Contains("d") || lower.Contains("lodz") || lower.Contains("dz")) return OddzialNFZ.lodz;
            if (lower.Contains("maop") || lower.Contains("krakw") || lower.Contains("krakow")) return OddzialNFZ.krakow;
            if (lower.Contains("mazow") || lower.Contains("warszaw")) return OddzialNFZ.warszawa;
            if (lower.Contains("wielk") || lower.Contains("pozna")) return OddzialNFZ.poznan;
            if (lower.Contains("opol")) return OddzialNFZ.opole;
            if (lower.Contains("podkar") || lower.Contains("rzeszw") || lower.Contains("rzeszow")) return OddzialNFZ.rzeszow;
            if (lower.Contains("podla") || lower.Contains("biays") || lower.Contains("bialys")) return OddzialNFZ.bialystok;
            if (lower.Contains("pomor") || lower.Contains("gda") || lower.Contains("gdan")) return OddzialNFZ.gdansk;
            if (lower.Contains("ls") || lower.Contains("slas") || lower.Contains("katow")) return OddzialNFZ.katowice;
            if (lower.Contains("wit") || lower.Contains("swiet") || lower.Contains("kielc")) return OddzialNFZ.kielce;
            if (lower.Contains("warmi") || lower.Contains("olszt")) return OddzialNFZ.olsztyn;
            if (lower.Contains("zacho") || lower.Contains("szcze")) return OddzialNFZ.szczecin;
            if (lower.Contains("dolno") || lower.Contains("wroc") || lower.Contains("wrocl")) 
                return (OddzialNFZ)Enum.Parse(typeof(OddzialNFZ), "wroclaw");
            if (lower.Contains("kujaw") || lower.Contains("bydgo")) 
                return (OddzialNFZ)Enum.Parse(typeof(OddzialNFZ), "bydgoszcz");
            return null;
        }

        internal static Panstwo? ResolvePanstwo(IUchwyt uchwyt, string inputName, out string platnikCitizenshipName)
        {
            return ZusCitizenshipManager.Instance.ResolvePanstwo(uchwyt, inputName, out platnikCitizenshipName);
        }

        internal static Obywatelstwo? ResolveObywatelstwo(IUchwyt uchwyt, string inputName)
        {
            return ZusCitizenshipManager.Instance.ResolveObywatelstwo(uchwyt, inputName);
        }

        private static string NormalizeUsFullName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            string norm = NormalizeForeignLetters(input).ToLowerInvariant();
            norm = norm.Replace("ą", "a").Replace("ć", "c").Replace("ę", "e")
                       .Replace("ł", "l").Replace("ń", "n").Replace("ó", "o")
                       .Replace("ś", "s").Replace("ź", "z").Replace("ż", "z");
            norm = norm.Replace("pierwszy", "1").Replace("drugi", "2").Replace("trzeci", "3");
            norm = System.Text.RegularExpressions.Regex.Replace(norm, @"[\-\/\.\,\'\(\)""\]]", " ");
            norm = System.Text.RegularExpressions.Regex.Replace(norm, @"\s+", " ").Trim();
            return norm;
        }

        private static string NormalizeUsToken(string word)
        {
            if (string.IsNullOrWhiteSpace(word)) return "";
            string w = NormalizeForeignLetters(word).ToLowerInvariant();
            w = w.Replace("ą", "a").Replace("ć", "c").Replace("ę", "e")
                 .Replace("ł", "l").Replace("ń", "n").Replace("ó", "o")
                 .Replace("ś", "s").Replace("ź", "z").Replace("ż", "z");
            w = w.Replace("pierwszy", "1").Replace("drugi", "2").Replace("trzeci", "3");
            return w;
        }

        private static HashSet<string> GetUsTokenSet(string input)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(input)) return set;

            var stopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "urzad", "skarbowy", "w", "we", "us", "podatkowy", "izba", "dla", "sekcja"
            };

            var words = input.Split(new[] { ' ', '-', '/', '.', ',', '"', '\'', '(', ')', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var rawWord in words)
            {
                string norm = NormalizeUsToken(rawWord);
                if (string.IsNullOrWhiteSpace(norm) || stopWords.Contains(norm))
                    continue;
                set.Add(norm);
            }
            return set;
        }

        private static string CleanAndStemUs(string input)
        {
            return NormalizeUsFullName(input);
        }

        internal static OrganPodatkowy? ResolveOrganPodatkowy(IUchwyt uchwyt, string usName)
        {
            if (string.IsNullOrWhiteSpace(usName)) return null;

            string normTargetName = NormalizeUsFullName(usName);
            var targetTokens = GetUsTokenSet(usName);
            if (string.IsNullOrEmpty(normTargetName) && targetTokens.Count == 0) return null;

            var allOrgany = uchwyt.Podmioty().Dane.WszystkieOrganyPodatkowe().AsEnumerable();

            // 1. Direct Normalized Exact Match on Local Database
            var exactLocal = allOrgany.FirstOrDefault(o => o.Firma != null && 
                (NormalizeUsFullName(o.Firma.Nazwa) == normTargetName || NormalizeUsFullName(o.NazwaSkrocona) == normTargetName));

            if (exactLocal != null)
            {
                LogNexo.Informacja($"[Import 2.0] Direct exact match found for US '{usName}': {exactLocal.Firma.Nazwa}");
                return exactLocal.Firma.OrganPodatkowy;
            }

            // 2. Token Set Matching on Local Database
            if (targetTokens.Count > 0)
            {
                OrganPodatkowy? bestLocal = null;
                int bestScore = 0;

                foreach (var o in allOrgany)
                {
                    if (o.Firma == null) continue;
                    var candTokens = GetUsTokenSet(o.Firma.Nazwa);
                    candTokens.UnionWith(GetUsTokenSet(o.NazwaSkrocona));

                    int matchCount = targetTokens.Count(t => candTokens.Contains(t));
                    // All target tokens must be present in candidate (or maximum match count)
                    if (matchCount > bestScore && matchCount == targetTokens.Count)
                    {
                        bestScore = matchCount;
                        bestLocal = o.Firma.OrganPodatkowy;
                    }
                }

                if (bestLocal != null)
                {
                    LogNexo.Informacja($"[Import 2.0] Token set match ({bestScore}/{targetTokens.Count} tokens) found for US '{usName}': {bestLocal.Firma?.Nazwa}");
                    return bestLocal;
                }
            }

            // 3. Search Nexo Master Template Dictionary (UrzedySkarbowe)
            try
            {
                var podmiotyMgr = uchwyt.Podmioty();
                
                string? kod = null;
                string? templateNazwa = null;
                
                using (var bo = podmiotyMgr.UtworzPracownika())
                {
                    var getUow = (InsERT.Mox.BusinessObjects.IGetUnitOfWork)bo;
                    var scope = (InsERT.Mox.Runtime.IInjectionScope)getUow.UnitOfWork;
                    var context = (ModelDanychContainer)scope.ScopedContainer.GetObject(typeof(ModelDanychContainer));
                    
                    var allTemplates = context.UrzedySkarbowe.AsEnumerable();

                    // Direct Exact Match on Templates
                    var exactTemplate = allTemplates.FirstOrDefault(t => t.Nazwa != null && NormalizeUsFullName(t.Nazwa) == normTargetName);
                    if (exactTemplate != null)
                    {
                        kod = exactTemplate.Kod;
                        templateNazwa = exactTemplate.Nazwa;
                    }
                    else if (targetTokens.Count > 0)
                    {
                        // Token Set Match on Templates
                        InsERT.Moria.ModelDanych.UrzadSkarbowy? bestTemplate = null;
                        int bestScore = 0;

                        foreach (var t in allTemplates)
                        {
                            if (string.IsNullOrWhiteSpace(t.Nazwa)) continue;
                            var candTokens = GetUsTokenSet(t.Nazwa);
                            int matchCount = targetTokens.Count(tok => candTokens.Contains(tok));
                            if (matchCount > bestScore && matchCount == targetTokens.Count)
                            {
                                bestScore = matchCount;
                                bestTemplate = t;
                            }
                        }

                        if (bestTemplate != null)
                        {
                            kod = bestTemplate.Kod;
                            templateNazwa = bestTemplate.Nazwa;
                        }
                    }
                }

                if (kod != null)
                {
                    using (var organBO = podmiotyMgr.UtworzOrganPodatkowy())
                    {
                        var organ = (IOrganPodatkowy)organBO;
                        if (organ.WypelnijDlaKodu(kod))
                        {
                            organBO.AutoSymbol();
                            if (organBO.Zapisz())
                            {
                                LogNexo.Informacja($"[Import 2.0] Utworzono brakujący urząd skarbowy z szablonu Nexo: {templateNazwa} (Kod: {kod})");
                                return organBO.Dane.Firma.OrganPodatkowy;
                            }
                            else
                            {
                                var bledy = uchwyt.PodajBledy(organBO);
                                var msg = string.Join("; ", bledy.Select(x => x.Tresc));
                                LogNexo.Informacja($"[Import 2.0] Błąd zapisu urzędu skarbowego {templateNazwa}: {msg}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogNexo.Blad(ex);
            }

            return null;
        }

        private static string ZapytajOUzytkownikaODrugieImie(string nazwisko, string imie, string opcja1, string opcja2)
        {
            string wybraneDrugieImie = $"{opcja1} {opcja2}";

            try
            {
                if (System.Windows.Application.Current != null && System.Windows.Application.Current.Dispatcher != null)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        var window = new System.Windows.Window
                        {
                            Title = "Wybr drugiego imienia (Import)",
                            Width = 520,
                            MinWidth = 480,
                            MinHeight = 320,
                            SizeToContent = System.Windows.SizeToContent.Height,
                            WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                            ResizeMode = System.Windows.ResizeMode.CanResizeWithGrip,
                            Topmost = true
                        };

                        var grid = new System.Windows.Controls.Grid { Margin = new System.Windows.Thickness(15) };
                        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
                        grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });

                        var scrollViewer = new System.Windows.Controls.ScrollViewer
                        {
                            MaxHeight = 250,
                            VerticalScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility.Auto,
                            Margin = new System.Windows.Thickness(0, 0, 0, 15)
                        };

                        var textBlock = new System.Windows.Controls.TextBlock
                        {
                            Text = $"Dla pracownika: {nazwisko} {imie}\nWykryto rozbieno w drugim imieniu:\n" +
                                   $"- Wycignite z pola 'Imi': \"{opcja1}\"\n" +
                                   $"- Wpisane w polu 'Drugie imi': \"{opcja2}\"\n\nWybierz, ktre drugie imi zastosowa:",
                            TextWrapping = System.Windows.TextWrapping.Wrap
                        };
                        scrollViewer.Content = textBlock;
                        System.Windows.Controls.Grid.SetRow(scrollViewer, 0);

                        var stack = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(0, 5, 0, 5) };
                        System.Windows.Controls.Grid.SetRow(stack, 1);

                        var btn1 = new System.Windows.Controls.Button { Content = $"Uyj: {opcja1} (z pola Imi)", Margin = new System.Windows.Thickness(0, 4, 0, 4), Height = 35, Padding = new System.Windows.Thickness(5) };
                        var btn2 = new System.Windows.Controls.Button { Content = $"Uyj: {opcja2} (z pola Drugie imi)", Margin = new System.Windows.Thickness(0, 4, 0, 4), Height = 35, Padding = new System.Windows.Thickness(5) };
                        var btn3 = new System.Windows.Controls.Button { Content = $"Uyj obu: {opcja1} {opcja2}", Margin = new System.Windows.Thickness(0, 4, 0, 4), Height = 35, Padding = new System.Windows.Thickness(5) };

                        btn1.Click += (s, e) => { wybraneDrugieImie = opcja1; window.DialogResult = true; window.Close(); };
                        btn2.Click += (s, e) => { wybraneDrugieImie = opcja2; window.DialogResult = true; window.Close(); };
                        btn3.Click += (s, e) => { wybraneDrugieImie = $"{opcja1} {opcja2}"; window.DialogResult = true; window.Close(); };

                        stack.Children.Add(btn1);
                        stack.Children.Add(btn2);
                        stack.Children.Add(btn3);

                        grid.Children.Add(scrollViewer);
                        grid.Children.Add(stack);

                        window.Content = grid;
                        window.ShowDialog();
                    });
                }
            }
            catch (Exception ex)
            {
                LogNexo.Informacja($"[Import 2.0] Nie mona wywietli okna wyboru drugiego imienia: {ex.Message}. Uyto obu wartoci: {wybraneDrugieImie}");
            }

            return wybraneDrugieImie;
        }

        private void PokazPodsumowanieImportu(IKontekstOperacji kontekstOperacji, List<PracownikImportRow> rows)
        {
            var zaimportowani = rows.Where(r => r.Status.StartsWith("Zaimportowano") || r.Status == "Gotowy" || r.Status == "Gotowy (Cudzoziemiec)").ToList();
            var pominieci = rows.Where(r => r.Status.StartsWith("Pominity") || r.Status.StartsWith("Pomini")).ToList();
            var bledy = rows.Where(r => r.Status.StartsWith("Bd") || r.Status.StartsWith("Bl") || r.Status == "Wyjtek" || r.Ostrzezenia.Contains("bd") || r.Ostrzezenia.Contains("Bd")).ToList();
            var zOstrzezeniami = rows.Where(r => !string.IsNullOrEmpty(r.Ostrzezenia) && !pominieci.Contains(r) && !bledy.Contains(r)).ToList();

            var sb = new StringBuilder();
            sb.AppendLine("Import 2.0 (Sfera Native) zakoczony.");
            sb.AppendLine($"cznie wierszy: {rows.Count}");
            sb.AppendLine($"Zaimportowano: {zaimportowani.Count}");
            sb.AppendLine($"Pominito (ju istniej): {pominieci.Count}");
            if (bledy.Count > 0)
                sb.AppendLine($"Bdy: {bledy.Count}");
            if (zOstrzezeniami.Count > 0)
                sb.AppendLine($"Z ostrzeeniami: {zOstrzezeniami.Count}");

            if (pominieci.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("--- Pominici pracownicy (ju istniej) ---");
                foreach (var p in pominieci)
                {
                    sb.AppendLine($"   {p.Nazwisko} {p.Imie} | PESEL: {p.Pesel} | Umowa: {p.OpisUmowy}");
                }
            }

            if (bledy.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("--- Bdy ---");
                foreach (var p in bledy)
                {
                    sb.AppendLine($"   {p.Nazwisko} {p.Imie} | {p.Status} | {p.Ostrzezenia}");
                }
            }

            if (zOstrzezeniami.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("--- Z ostrzeeniami ---");
                foreach (var p in zOstrzezeniami)
                {
                    sb.AppendLine($"   {p.Nazwisko} {p.Imie} | {p.Ostrzezenia}");
                }
            }

            try
            {
                Okna.PokazOknoZInformacja(kontekstOperacji.Uchwyt, sb.ToString());
            }
            catch
            {
                try
                {
                    System.Windows.MessageBox.Show(sb.ToString(), "Import 2.0 (Sfera Native) - Podsumowanie", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
                catch { }
            }

            LogNexo.Informacja($"[Import 2.0] Podsumowanie importu: {sb}");
        }

        private static (string ulica, string nrDomu, string nrLokalu) ParseSmartAddress(string rawUlica, string rawHouse, string rawCity)
        {
            var norm = AddressNormalizer.Normalize(
                rawStreet: rawUlica,
                rawHouseAndApt: rawHouse,
                rawCity: rawCity,
                rawPostalCode: null
            );
            return (norm.Ulica, norm.NrDomu, norm.NrLokalu);
        }

        private static bool ParseBankString(string rawBankStr, out string iban, out string currencySymbol, out string bankName)
        {
            iban = "";
            currencySymbol = "PLN";
            bankName = "Rachunek podstawowy";

            if (string.IsNullOrWhiteSpace(rawBankStr))
                return false;

            // Regex match for IBAN (optional PL prefix + 26 digits with optional spaces, or 26 digits)
            var match = System.Text.RegularExpressions.Regex.Match(rawBankStr, @"(PL\s*)?\d{2}(\s*\d{4}){6}|\d{26}", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (match.Success)
            {
                iban = match.Value.Replace(" ", "").ToUpper();
                if (iban.Length > 34)
                {
                    iban = iban.Substring(0, 34);
                }

                string remainder = (rawBankStr.Substring(0, match.Index) + rawBankStr.Substring(match.Index + match.Length)).Trim();
                remainder = remainder.Trim('\r', '\n', ' ', '-', ':', '|');

                var currMatch = System.Text.RegularExpressions.Regex.Match(remainder, @"^(PLN|EUR|USD|GBP|CHF)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                if (currMatch.Success)
                {
                    currencySymbol = currMatch.Value.ToUpper();
                    remainder = remainder.Substring(currMatch.Length).Trim('\r', '\n', ' ', '-', ':', '|');
                }

                if (!string.IsNullOrWhiteSpace(remainder))
                {
                    bankName = System.Text.RegularExpressions.Regex.Replace(remainder, @"[\r\n]+", " ").Trim();
                    if (bankName.Length > 255)
                    {
                        bankName = bankName.Substring(0, 255);
                    }
                }

                return true;
            }
            else
            {
                string[] bankParts = rawBankStr.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                iban = bankParts.Length >= 1 ? bankParts[0].Trim().Replace(" ", "").ToUpper() : "";
                if (iban.Length > 34)
                {
                    iban = iban.Substring(0, 34);
                }
                if (bankParts.Length >= 2)
                {
                    currencySymbol = bankParts[1].Trim().ToUpper();
                }
                if (bankParts.Length >= 3)
                {
                    bankName = bankParts[2].Trim();
                }

                return !string.IsNullOrWhiteSpace(iban);
            }
        }
    }

    public static class WpfWarningResolverDialog
    {
        public static DateTime? RozwiazDataUrodzenia(string nazwisko, string imie, DateTime peselDate, DateTime excelDate)
        {
            DateTime? chosen = null;
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                var w = new System.Windows.Window { Title = "Rozbie?no?? daty urodzenia", Width=450, SizeToContent=System.Windows.SizeToContent.Height, WindowStartupLocation=System.Windows.WindowStartupLocation.CenterScreen };
                var s = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(10) };
                s.Children.Add(new System.Windows.Controls.TextBlock { Text = $"Dla: {nazwisko} {imie}\nPESEL: {peselDate:yyyy-MM-dd}\nExcel: {excelDate:yyyy-MM-dd}", Margin=new System.Windows.Thickness(0,0,0,10) });
                var b1 = new System.Windows.Controls.Button { Content = $"U?yj z PESEL: {peselDate:yyyy-MM-dd}", Margin = new System.Windows.Thickness(2) };
                var b2 = new System.Windows.Controls.Button { Content = $"U?yj z Excela: {excelDate:yyyy-MM-dd}", Margin = new System.Windows.Thickness(2) };
                b1.Click += (se, ev) => { chosen = peselDate; w.DialogResult = true; w.Close(); };
                b2.Click += (se, ev) => { chosen = excelDate; w.DialogResult = true; w.Close(); };
                s.Children.Add(b1);
                s.Children.Add(b2);
                w.Content = s;
                w.ShowDialog();
            });
            return chosen;
        }

        public static string? RozwiazPlec(string nazwisko, string imie, string peselPlec, string excelPlec)
        {
            string? chosen = null;
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                var w = new System.Windows.Window { Title = "Rozbie?no?? p?ci", Width=450, SizeToContent=System.Windows.SizeToContent.Height, WindowStartupLocation=System.Windows.WindowStartupLocation.CenterScreen };
                var s = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(10) };
                s.Children.Add(new System.Windows.Controls.TextBlock { Text = $"Dla: {nazwisko} {imie}\nPESEL: {peselPlec}\nExcel: {excelPlec}", Margin=new System.Windows.Thickness(0,0,0,10) });
                var b1 = new System.Windows.Controls.Button { Content = $"U?yj z PESEL: {peselPlec}", Margin = new System.Windows.Thickness(2) };
                var b2 = new System.Windows.Controls.Button { Content = $"U?yj z Excela: {excelPlec}", Margin = new System.Windows.Thickness(2) };
                b1.Click += (se, ev) => { chosen = peselPlec; w.DialogResult = true; w.Close(); };
                b2.Click += (se, ev) => { chosen = excelPlec; w.DialogResult = true; w.Close(); };
                s.Children.Add(b1);
                s.Children.Add(b2);
                w.Content = s;
                w.ShowDialog();
            });
            return chosen;
        }

        public static InsERT.Moria.ModelDanych.OrganPodatkowy? RozwiazBrakUrzeduSkarbowego(string nazwisko, string imie, string usZExcela, IEnumerable<InsERT.Moria.ModelDanych.OrganPodatkowy> wszystkieUrzedy)
        {
            InsERT.Moria.ModelDanych.OrganPodatkowy? chosen = null;
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                var w = new System.Windows.Window { Title = "Brakuj?cy lub nierozpoznany Urz?d Skarbowy", Width=450, SizeToContent=System.Windows.SizeToContent.Height, WindowStartupLocation=System.Windows.WindowStartupLocation.CenterScreen };
                var s = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(10) };
                s.Children.Add(new System.Windows.Controls.TextBlock { Text = $"Dla: {nazwisko} {imie}\nNie znaleziono urz?du z Excela: {usZExcela}\nWybierz z listy:", Margin=new System.Windows.Thickness(0,0,0,10) });
                
                var cb = new System.Windows.Controls.ComboBox { Margin = new System.Windows.Thickness(0,0,0,10), DisplayMemberPath = "" };
                var list = wszystkieUrzedy.OrderBy(u => u.ToString()).ToList();
                cb.ItemsSource = list;
                
                var b1 = new System.Windows.Controls.Button { Content = "Zatwierd? wyb?r", Margin = new System.Windows.Thickness(2) };
                var b2 = new System.Windows.Controls.Button { Content = "Pomi? urz?d (zostaw puste)", Margin = new System.Windows.Thickness(2) };
                b1.Click += (se, ev) => { chosen = cb.SelectedItem as InsERT.Moria.ModelDanych.OrganPodatkowy; w.DialogResult = true; w.Close(); };
                b2.Click += (se, ev) => { chosen = null; w.DialogResult = true; w.Close(); };
                s.Children.Add(cb);
                s.Children.Add(b1);
                s.Children.Add(b2);
                w.Content = s;
                w.ShowDialog();
            });
            return chosen;
        }

        public static DecyzjaIstniejacyPracownik RozwiazKonfliktPracownika(string nazwisko, string imie, string identifierType, string identifierVal)
        {
            DecyzjaIstniejacyPracownik chosen = DecyzjaIstniejacyPracownik.AktualizujKartotekeIDodajUmowe;
            System.Windows.Application.Current?.Dispatcher.Invoke(() =>
            {
                var w = new System.Windows.Window { Title = "Wykryto istniej?cego pracownika", Width=450, SizeToContent=System.Windows.SizeToContent.Height, WindowStartupLocation=System.Windows.WindowStartupLocation.CenterScreen };
                var s = new System.Windows.Controls.StackPanel { Margin = new System.Windows.Thickness(10) };
                s.Children.Add(new System.Windows.Controls.TextBlock { Text = $"Dla: {nazwisko} {imie}\nIstnieje ju? pracownik z {identifierType}: {identifierVal}.\nWybierz akcj?:", Margin=new System.Windows.Thickness(0,0,0,10) });
                
                var b1 = new System.Windows.Controls.Button { Content = "Aktualizuj kartotek? i dodaj umow?", Margin = new System.Windows.Thickness(2) };
                var b2 = new System.Windows.Controls.Button { Content = "Tylko dodaj umow? (nie zmieniaj danych)", Margin = new System.Windows.Thickness(2) };
                var b3 = new System.Windows.Controls.Button { Content = "Nadpisz istniej?c? umow? z t? sam? dat?", Margin = new System.Windows.Thickness(2) };
                var b4 = new System.Windows.Controls.Button { Content = "Pomi? wiersz ca?kowicie", Margin = new System.Windows.Thickness(2) };

                b1.Click += (se, ev) => { chosen = DecyzjaIstniejacyPracownik.AktualizujKartotekeIDodajUmowe; w.DialogResult = true; w.Close(); };
                b2.Click += (se, ev) => { chosen = DecyzjaIstniejacyPracownik.TylkoDodajUmowe; w.DialogResult = true; w.Close(); };
                b3.Click += (se, ev) => { chosen = DecyzjaIstniejacyPracownik.NadpiszUmowe; w.DialogResult = true; w.Close(); };
                b4.Click += (se, ev) => { chosen = DecyzjaIstniejacyPracownik.PominWiersz; w.DialogResult = true; w.Close(); };

                s.Children.Add(b1);
                s.Children.Add(b2);
                s.Children.Add(b3);
                s.Children.Add(b4);
                w.Content = s;
                w.ShowDialog();
            });
            return chosen;
        }
    }
}










