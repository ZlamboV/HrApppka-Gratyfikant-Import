#pragma warning disable CA1416
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using InsERT.Moria.Flagi;
using InsERT.Moria.Kadry.Duze;
using InsERT.Moria.Klienci;
using InsERT.Moria.Listy;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Rozszerzanie.Operacje;
using InsERT.Moria.Sfera;
using InsERT.Moria.DeklaracjeZUS;
using Microsoft.Win32;
using MiniExcelLibs;

namespace HrAppka_Import_Pracowników
{
    internal class OperacjaImportuPracownikow : OperacjaNaLiscieDanych<Pracownik, int>
    {
        protected override string[]? SciezkaWMenu => null;

        public override string Nazwa => "Importuj z HrAppka";

        protected override bool SprawdzCzyMoznaWykonac(IReadOnlyCollection<int> identyfikatoryWybranychElementow, IKontekstListyDanych kontekstListyDanych, IKontekstOperacji kontekstOperacji)
        {
            return true;
        }

        protected override void Wykonaj(IReadOnlyCollection<int> identyfikatoryWybranychElementow, IKontekstListyDanych kontekstListyDanych, IKontekstOperacji kontekstOperacji)
        {
            if (PodajOSciezkeDoPliku(out string sciezka))
            {
                List<PracownikImportRow> rows;
                try
                {
                    // odczyt listy jakichś elementów
                    LogNexo.Informacja("Rozpoczęto pobieranie danych");
                    rows = WczytajExcel(sciezka);
                    // koniec operacji
                    LogNexo.Informacja("Operacja zakończona");
                }
                catch (ArgumentException argException)
                {
                    LogNexo.Blad(argException);
                    Okna.PokazOknoZBledem(kontekstOperacji.Uchwyt, $"Błąd argumentu: {argException.Message}");
                    return;
                }
                catch (Exception ex)
                {
                    LogNexo.Blad(ex);
                    Okna.PokazOknoZBledem(kontekstOperacji.Uchwyt, $"Błąd odczytu pliku Excel: {ex.Message}");
                    return;
                }

                if (rows == null || !rows.Any())
                {
                    Okna.PokazOknoZBledem(kontekstOperacji.Uchwyt, "Wskazany plik nie zawiera danych w wymaganym formacie.");
                    return;
                }

                ImportujDanePracownikow(kontekstOperacji, rows);
            }
        }

        private bool PodajOSciezkeDoPliku(out string sciezka)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Pliki Excel (*.xlsx)|*.xlsx";
            dialog.Title = "Wybierz plik z wyeksportowanymi danymi HrAppka";
            dialog.CheckFileExists = true;

            var result = dialog.ShowDialog();
            if (result == true)
            {
                sciezka = dialog.FileName;
                return true;
            }
            sciezka = "";
            return false;
        }

        private List<PracownikImportRow> WczytajExcel(string sciezka)
        {
            var rows = new List<PracownikImportRow>();
            var excelData = MiniExcel.Query(sciezka, useHeaderRow: true).Cast<IDictionary<string, object>>().ToList();
            
            foreach (var d in excelData)
            {
                var row = new PracownikImportRow();
                row.Imie = GetVal(d, "Pracownik - Imię");
                row.DrugieImie = GetVal(d, "Pracownik - Drugie imię");
                row.Nazwisko = GetVal(d, "Pracownik - Nazwisko");
                row.PlecExcel = GetVal(d, "Pracownik - Płeć");
                row.BirthDateExcel = GetVal(d, "Pracownik - Data urodzenia");
                row.Pesel = GetVal(d, "Pracownik - Numery identyfikacyjne - PESEL");
                row.Obywatelstwo = GetVal(d, "Pracownik - Obywatelstwo");
                row.Paszport = GetVal(d, "Pracownik - Numery identyfikacyjne - Paszport");
                row.PaszportWydany = GetVal(d, "Pracownik - Numery identyfikacyjne - Paszport - Data wydania");
                row.PaszportWazny = GetVal(d, "Pracownik - Numery identyfikacyjne - Paszport - Data ważności");
                
                row.Ulica = GetVal(d, "Pracownik - Adres Zamieszkania - Ulica");
                row.NrDomu = GetVal(d, "Pracownik - Adres Zamieszkania - Numer domu / mieszkania");
                row.Miejscowosc = GetVal(d, "Pracownik - Adres Zamieszkania - Miejscowość");
                row.KodPocztowy = GetVal(d, "Pracownik - Adres Zamieszkania - Kod pocztowy");
                row.Gmina = GetVal(d, "Pracownik - Adres Zamieszkania - Gmina");
                row.Powiat = GetVal(d, "Pracownik - Adres Zamieszkania - Powiat");
                row.Wojewodztwo = GetVal(d, "Pracownik - Adres Zamieszkania - Województwo");
                row.Kraj = GetVal(d, "Pracownik - Adres Zamieszkania - Kraj");
                
                row.Telefon = GetVal(d, "Pracownik - Dane kontaktowe - Telefon");
                row.Email = GetVal(d, "Pracownik - Dane kontaktowe - E-mail");
                
                row.BankStr = GetVal(d, "Pracownik - Rachunek bankowy");
                row.NfzStr = GetVal(d, "Pracownik - Oddział NFZ");
                string us = GetVal(d, "Pracownik - Urząd skarbowy");
                if (string.IsNullOrWhiteSpace(us))
                {
                    us = GetVal(d, "Urząd skarbowy");
                }
                row.UrzadSkarbowy = us;
                
                row.TypUmowy = GetVal(d, "Typ");
                row.DataRozpoczecia = GetVal(d, "Data rozpoczęcia");
                row.DataZakonczenia = GetVal(d, "Data zakończenia");
                row.DataZawarcia = GetVal(d, "Data zawarcia");
                row.DzialNazwa = GetVal(d, "Dział/Klient - Nazwa");
                row.KodZawodu = GetVal(d, "Pracownik - Bieżąca umowa - Stanowisko - Kod zawodu");
                
                row.JestStudentem = GetVal(d, "Pracownik - Dane do rozliczeń - Jestem studentem lub uczniem i nie przekroczyłem 26 roku życia");
                
                if (!string.IsNullOrWhiteSpace(row.Imie) || !string.IsNullOrWhiteSpace(row.Nazwisko) || !string.IsNullOrWhiteSpace(row.Pesel))
                {
                    rows.Add(row);
                }
            }
            return rows;
        }

        private string GetVal(IDictionary<string, object> dict, string key)
        {
            if (dict.TryGetValue(key, out var val) && val != null)
            {
                return val.ToString()?.Trim() ?? "";
            }
            var keyNormalized = NormalizeKey(key);
            foreach (var k in dict.Keys)
            {
                if (NormalizeKey(k) == keyNormalized)
                {
                    return dict[k]?.ToString()?.Trim() ?? "";
                }
            }
            return "";
        }

        private string NormalizeKey(string key)
        {
            return key.ToLower()
                      .Replace("ą", "a").Replace("ć", "c").Replace("ę", "e")
                      .Replace("ł", "l").Replace("ń", "n").Replace("ó", "o")
                      .Replace("ś", "s").Replace("ź", "z").Replace("ż", "z")
                      .Replace(" ", "").Replace("-", "").Replace("/", "");
        }

        private void ImportujDanePracownikow(IKontekstOperacji kontekstOperacji, List<PracownikImportRow> rows)
        {
            var okno = kontekstOperacji.Uchwyt.PodajObiektTypu<IOknoOperacjiZbiorczej>();
            okno.Tytul = "Import pracowników z HrAppka";
            okno.TekstPrzyciskuWykonaj = "Importuj";
            
            // Pre-analyze duplicates and validate PESEL before displaying
            PreAnalyzeRows(kontekstOperacji.Uchwyt, rows);

            okno.Pokaz(rows, x => ZaimportujRow(kontekstOperacji, x));
        }

        private void PreAnalyzeRows(IUchwyt uchwyt, List<PracownikImportRow> rows)
        {
            var podmioty = uchwyt.Podmioty();
            var existingPesels = podmioty.Dane.WszyscyPracownicy()
                .Select(p => p.Osoba.PESEL)
                .Where(pesel => !string.IsNullOrEmpty(pesel))
                .ToHashSet();

            var existingPassports = podmioty.Dane.WszyscyPracownicy()
                .Select(p => p.Osoba.SeriaINumerDokumentuTozsamosci)
                .Where(passport => !string.IsNullOrEmpty(passport))
                .ToHashSet();

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
                            row.Ostrzezenia += "Pracownik o tym Paszporcie już istnieje. ";
                            row.Status = "Pominięty (Już istnieje)";
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
                        if (!string.IsNullOrWhiteSpace(row.BirthDateExcel))
                        {
                            if (DateTime.TryParse(row.BirthDateExcel, out DateTime dtExcel))
                            {
                                if (dtExcel.Date != birthDate.Value.Date)
                                {
                                    row.Ostrzezenia += $"Niezgodność daty urodzenia z PESEL (Excel: {dtExcel:yyyy-MM-dd}, PESEL: {birthDate.Value:yyyy-MM-dd}). ";
                                }
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(row.PlecExcel))
                        {
                            if (!string.Equals(row.PlecExcel, gender, StringComparison.OrdinalIgnoreCase))
                            {
                                row.Ostrzezenia += $"Niezgodność płci z PESEL (Excel: {row.PlecExcel}, PESEL: {gender}). ";
                            }
                        }
                    }

                    if (existingPesels.Contains(row.Pesel))
                    {
                        row.Ostrzezenia += "Pracownik o tym PESEL już istnieje. ";
                        row.Status = "Pominięty (Już istnieje)";
                    }
                }

                if (!string.IsNullOrWhiteSpace(row.BankStr))
                {
                    string[] bankParts = row.BankStr.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                    if (bankParts.Length < 3)
                    {
                        row.Ostrzezenia += "Niepełne dane rachunku bankowego w pliku. ";
                    }
                }
            }
        }

        private WynikOperacji ZaimportujRow(IKontekstOperacji kontekstOperacji, PracownikImportRow item)
        {
            var uchwyt = kontekstOperacji.Uchwyt;
            LogNexo.Informacja($"Rozpoczęto import pracownika: {item.Imie} {item.Nazwisko}");
            
            if (item.Status.StartsWith("Pominięty"))
            {
                LogNexo.Informacja($"Pominięto pracownika: {item.Imie} {item.Nazwisko} - pracownik już istnieje.");
                return WynikOperacji.ZakonczonaPowodzeniem("Pominięto - pracownik już istnieje.");
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
                Podmiot employeePodmiot;

                IPodmiot pracownik;
                Podmiot existingOsoba = null;

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

                if (existingOsoba != null)
                {
                    LogNexo.Informacja($"Znaleziono istniejącą osobę w Nexo: ID={existingOsoba.Id}. Nastąpi aktualizacja danych.");
                    if (existingOsoba.Osoba.Pracownik != null)
                    {
                        pracownik = podmiotyMgr.Znajdz(existingOsoba);
                    }
                    else
                    {
                        pracownik = podmiotyMgr.PrzeksztalcOsobeNaPracownika(existingOsoba);
                    }
                }
                else
                {
                    LogNexo.Informacja("Nie znaleziono osoby o podanym PESEL/Paszporcie. Tworzenie nowego pracownika.");
                    pracownik = podmiotyMgr.UtworzPracownika();
                    pracownik.AutoSymbol();
                }

                using (pracownik)
                {
                    var podmiot = pracownik.Dane;
                    var osoba = podmiot.Osoba;

                    // Tax Office (Urząd Skarbowy) - Hybrid Resolution
                    if (!string.IsNullOrWhiteSpace(item.UrzadSkarbowy))
                    {
                        var organ = ResolveOrganPodatkowy(uchwyt, item.UrzadSkarbowy);
                        if (organ != null)
                        {
                            podmiot.OrganPodatkowy = organ;
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
                                {
                                    podmiot.Uwagi = podmiot.Uwagi.Substring(0, 256);
                                }
                            }
                            item.Ostrzezenia += $"Nie znaleziono urzędu skarbowego '{item.UrzadSkarbowy}' w bazie. Pozostawiono pustym, zapisano w uwagach. ";
                        }
                    }
                    else
                    {
                        podmiot.OrganPodatkowy = null;
                    }


                    // Personal details
                    osoba.Imie = item.Imie?.ToUpper().Trim();
                    osoba.DrugieImie = item.DrugieImie?.ToUpper().Trim();
                    osoba.Nazwisko = item.Nazwisko?.ToUpper().Trim();
                    osoba.PESEL = item.Pesel;
                    if (finalBirthDate.HasValue)
                        osoba.DataUrodzenia = finalBirthDate.Value;

                    if (!string.IsNullOrEmpty(finalGender))
                    {
                        osoba.Plec = (byte)GetPlec(finalGender);
                    }

                    if (!string.IsNullOrWhiteSpace(item.Obywatelstwo))
                    {
                        var obywatelstwo = ResolveObywatelstwo(uchwyt, item.Obywatelstwo);
                        if (obywatelstwo != null)
                        {
                            osoba.Pracownik.Obywatelstwo = obywatelstwo;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(item.Miejscowosc) && !string.IsNullOrWhiteSpace(item.NrDomu))
                    {
                        var adres = podmiot.AdresPodstawowy;
                        if (adres == null)
                        {
                            adres = pracownik.DodajAdres(uchwyt.TypyAdresu().DaneDomyslne.Glowny);
                        }
                        adres.Nazwa = "Adres zamieszkania";
                        adres.Szczegoly.Ulica = item.Ulica;
                        adres.Szczegoly.NrDomu = item.NrDomu;
                        adres.Szczegoly.Miejscowosc = item.Miejscowosc;
                        adres.Szczegoly.KodPocztowy = item.KodPocztowy;
                        adres.Szczegoly.Poczta = item.Miejscowosc;
                        adres.Panstwo = ZnajdzPanstwo(uchwyt, item.Kraj);

                        var woj = ZnajdzWojewodztwo(uchwyt, item.Wojewodztwo);
                        if (woj != null) adres.Szczegoly.Wojewodztwo = woj;

                        var pow = ZnajdzPowiat(uchwyt, woj, item.Powiat);
                        if (pow != null)
                        {
                            var gmina = ZnajdzGmine(uchwyt, pow, item.Gmina);
                            if (gmina != null) adres.Szczegoly.Gmina = gmina;
                        }
                        else
                        {
                            var gmina = ZnajdzGmine(uchwyt, null, item.Gmina);
                            if (gmina != null) adres.Szczegoly.Gmina = gmina;
                        }
                    }

                    // Contacts (avoiding duplicates when editing existing entity)
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
                        var sb = new System.Text.StringBuilder();
                        for (int i = 0; i < item.Telefon.Length; i++)
                        {
                            char c = item.Telefon[i];
                            if (char.IsDigit(c))
                            {
                                sb.Append(c);
                            }
                            else if (c == '+' && sb.Length == 0)
                            {
                                sb.Append(c);
                            }
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

                    // Passport identity document details
                    if (!string.IsNullOrWhiteSpace(item.Paszport))
                    {
                        osoba.TypDokumentuTozsamosci = (byte)InsERT.Moria.Klienci.TypDokumentuTozsamosci.Paszport;
                        osoba.SeriaINumerDokumentuTozsamosci = item.Paszport;
                        if (DateTime.TryParse(item.PaszportWydany, out DateTime wydany))
                            osoba.DataWystawieniaDokumentuTozsamosci = wydany;
                    }

                    // Gratyfikant-specific fields on pracownik.Dane.Osoba.Pracownik.PracownikGr
                    var pracownikGr = osoba.Pracownik.PracownikGr;

                    DateTime contractStart = DateTime.Today;
                    if (!string.IsNullOrWhiteSpace(item.DataRozpoczecia) && DateTime.TryParse(item.DataRozpoczecia, out DateTime parsedStart))
                    {
                        contractStart = parsedStart;
                    }

                    pracownikGr.DaneWdrozenioweGr.MiesiacPierwszejWyplaty = new DateTime(contractStart.Year, contractStart.Month, 1);
                    pracownikGr.DataPrzystapieniaDoNFZ = contractStart;
                    pracownikGr.Wyksztalcenie = (byte)0; // Default education level cast to byte

                    if (!string.IsNullOrWhiteSpace(item.NfzStr))
                    {
                        var nfz = MappedNFZ(item.NfzStr);
                        if (nfz.HasValue)
                        {
                            pracownikGr.OddzialNFZ = (byte)nfz.Value;
                        }
                    }

                    // Bank Account
                    if (!string.IsNullOrWhiteSpace(item.BankStr))
                    {
                        string[] bankParts = item.BankStr.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
                        if (bankParts.Length >= 3)
                        {
                            string iban = bankParts[0].Trim();
                            string currencySymbol = bankParts[1].Trim();
                            string bankName = bankParts[2].Trim();

                            var currency = uchwyt.PodajObiektTypu<InsERT.Moria.Waluty.IWaluty>().Dane.Wszystkie().FirstOrDefault(w => w.Symbol == currencySymbol)
                                           ?? uchwyt.PodajObiektTypu<InsERT.Moria.Waluty.IWaluty>().Dane.Wszystkie().FirstOrDefault(w => w.Symbol == "PLN");

                            var existingRachunek = podmiot.Rachunki.FirstOrDefault(r => r.Numer == iban);
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
                        }
                    }

                    if (!pracownik.Zapisz())
                    {
                        var bledy = uchwyt.PodajBledy(pracownik);
                        var msg = string.Join("; ", bledy.Select(x => $"[{x.Waznosc}] {x.Tresc}"));
                        
                        // Fallback: deep traversal of all UoW participants and data graph
                        if (string.IsNullOrWhiteSpace(msg))
                        {
                            try
                            {
                                var errorList = new System.Collections.Generic.List<string>();
                                var bo = (InsERT.Mox.BusinessObjects.IBusinessObject)pracownik;
                                var uow = ((InsERT.Mox.BusinessObjects.IGetUnitOfWork)pracownik).UnitOfWork;
                                
                                // Traverse ALL participants in the Unit of Work
                                foreach (var participant in uow.Participants.OfType<InsERT.Mox.BusinessObjects.IBusinessObject>())
                                {
                                    var dd = (InsERT.Mox.DataAccess.IGetDataDomain)participant;
                                    dd.DataDomain.TraverseData(participant.Data, (obj) =>
                                    {
                                        var errInfo = obj as InsERT.Mox.Validation.ITypedDataErrorInfo;
                                        if (errInfo != null)
                                        {
                                            foreach (var err in errInfo.Errors)
                                                errorList.Add($"[{obj.GetType().Name}] {err}");
                                            foreach (var memberErr in errInfo.MemberErrors)
                                                errorList.Add($"[{obj.GetType().Name}.{string.Join(", ", memberErr)}] {memberErr.Key}");
                                        }
                                        return true;
                                    });
                                }
                                
                                if (errorList.Any())
                                    msg = string.Join("; ", errorList);
                                else
                                    msg = "Zapisz() zwrócił false, brak błędów walidacji w UoW. Prawdopodobnie błąd wewnętrzny Sfera.";
                            }
                            catch (Exception diagEx)
                            {
                                msg = $"Zapisz() zwrócił false. Diagnostyka: {diagEx.Message}";
                            }
                        }

                        LogNexo.Informacja($"Błąd zapisu pracownika {item.Imie} {item.Nazwisko}: {msg}");
                        return WynikOperacji.ZakonczonaNiepowodzeniem($"Błąd zapisu pracownika: {msg}");
                    }

                    employeePodmiot = pracownik.Dane;
                }

                // Civil Contract
                if (!string.IsNullOrWhiteSpace(item.TypUmowy) && !string.IsNullOrWhiteSpace(item.DataRozpoczecia))
                {
                    if (!DateTime.TryParse(item.DataRozpoczecia, out DateTime startDate))
                    {
                        item.Ostrzezenia += "Niepoprawna data rozpoczęcia umowy. ";
                        return WynikOperacji.ZakonczonaPowodzeniem("Pominięto umowę - błędna data rozpoczęcia.");
                    }
                    DateTime? endDate = null;
                    if (!string.IsNullOrWhiteSpace(item.DataZakonczenia) && DateTime.TryParse(item.DataZakonczenia, out DateTime parsedEnd))
                    {
                        endDate = parsedEnd;
                    }
                    DateTime signDate = startDate;
                    if (!string.IsNullOrWhiteSpace(item.DataZawarcia) && DateTime.TryParse(item.DataZawarcia, out DateTime parsedSign))
                    {
                        signDate = parsedSign;
                    }

                    var umowaZlecenie = uchwyt.TypyUmowPracowniczych().Dane.Wszystkie().FirstOrDefault(x => x.Nazwa == "Umowa zlecenie");
                    if (umowaZlecenie == null)
                    {
                        return WynikOperacji.ZakonczonaNiepowodzeniem("Nie znaleziono typu umowy 'Umowa zlecenie' w bazie.");
                    }

                    using (var umowa = uchwyt.UmowyPracowniczeGr().Utworz(umowaZlecenie))
                    {
                        var baseUmowa = umowa.Dane;

                        // Link employee and period
                        baseUmowa.Pracownik = employeePodmiot.Osoba.Pracownik;
                        
                        if (baseUmowa.OkresObowiazywania == null)
                        {
                            baseUmowa.OkresObowiazywania = new OkresWymaganyLewostronnie();
                        }
                        baseUmowa.OkresObowiazywania.DataPoczatkowa = startDate;
                        baseUmowa.OkresObowiazywania.DataKoncowa = endDate;

                        // Set the contract relationship type explicitly
                        baseUmowa.RodzajUmowyPracowniczej = umowaZlecenie.DozwoloneRodzajeUmowyPracowniczej.FirstOrDefault(x => x.Nazwa == "Na czas określony");

                        // Assign the print template 'WzorzecTresciUmowyPracowniczej'
                        baseUmowa.WzorzecTresciUmowyPracowniczej = uchwyt.TresciWydrukuRTF().DaneDomyslne.UCPZlecenieStale;

                        // ZUS & PIT Settings
                        int age = PeselHelper.CalculateAge(finalBirthDate ?? DateTime.MinValue, startDate);
                        bool studentFlag = string.Equals(item.JestStudentem, "Tak", StringComparison.OrdinalIgnoreCase);
                        bool ageUnder26 = age < 26;

                        bool isStudentUnder26 = studentFlag && ageUnder26;
                        bool isZusSubject = !isStudentUnder26;

                        // ZUS Report
                        baseUmowa.RaportZUS = (byte)(isStudentUnder26 ? InsERT.Moria.Kadry.TypRaportuZUS.Brak : InsERT.Moria.Kadry.TypRaportuZUS.RCA);

                        // ZUS Contributions
                        baseUmowa.NaliczaSkladkiEmerytalne = isZusSubject;
                        baseUmowa.NaliczaSkladkiRentowe = isZusSubject;
                        baseUmowa.NaliczaSkladkiChorobowe = false;
                        baseUmowa.NaliczaSkladkiWypadkowe = isZusSubject;
                        baseUmowa.NaliczaSkladkiZdrowotne = isZusSubject;
                        baseUmowa.NaliczaSkladkiFP = isZusSubject;
                        baseUmowa.NaliczaSkladkiFGSP = isZusSubject;

                        // PIT Exemption (youth under 26 is automatically handled via null "Tak (wg typu umowy)")
                        baseUmowa.StosujZwolnienieZPITDo26Lat = null;

                        baseUmowa.DataSporzadzenia = signDate;



                        // Department (Dział)
                        var subDept = DepartmentResolver.ResolveDepartment(uchwyt, item.DzialNazwa);
                        if (subDept != null)
                        {
                            var st = new StanowiskoWUmowiePracowniczej();
                            baseUmowa.Stanowiska.Add(st);
                            st.DataOd = startDate;
                            st.JednostkaOrganizacyjna = subDept;
                            st.LP = 1;
                        }

                        // Occupation (Kod zawodu)
                        if (!string.IsNullOrWhiteSpace(item.KodZawodu))
                        {
                            var zawod = uchwyt.PodajObiektTypu<InsERT.Moria.Kadry.Duze.IZawody>().Dane.Wszystkie().FirstOrDefault(z => z.Kod == item.KodZawodu);
                            if (zawod != null)
                            {
                                baseUmowa.Zawod = zawod;
                            }
                        }

                        // ZUS Limits & Flags (only set limits if the corresponding contributions are being calculated)
                        if (baseUmowa.NaliczaSkladkiZdrowotne)
                        {
                            baseUmowa.NieOgraniczaZdrowotnego = false;
                        }
                        if (baseUmowa.NaliczaSkladkiChorobowe)
                        {
                            baseUmowa.NieOgraniczaChorobowego = false;
                        }
                        if (baseUmowa.NaliczaSkladkiFP)
                        {
                            baseUmowa.NaliczaSkladkiFPTylkoPowyzejMinimalnej = isZusSubject;
                        }

                        // PIT Advance & Relief
                        try
                        {
                            var targetZaliczka = (byte)(ageUnder26 
                                ? InsERT.Moria.Kadry.Duze.SposobPobieraniaZaliczkiNaPodatekNaUmowie.WgUstawienPracownika 
                                : InsERT.Moria.Kadry.Duze.SposobPobieraniaZaliczkiNaPodatekNaUmowie.NiePobierajDoKwotyWolnej);
                            if (baseUmowa.ZaliczkaNaPodatek != targetZaliczka)
                            {
                                baseUmowa.ZaliczkaNaPodatek = targetZaliczka;
                            }

                            var targetUlga = (byte)InsERT.Moria.Kadry.Duze.SposobOdliczaniaUlgiPodatkowejNaUmowie.WgUstawienPracownika;
                            if (baseUmowa.UlgaPodatkowa != targetUlga)
                            {
                                baseUmowa.UlgaPodatkowa = targetUlga;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogNexo.Informacja($"Ignorowano błąd ustawiania PIT: {ex.Message}");
                        }

                        // Payment Schedule
                        try
                        {
                            var targetWyplata = (byte)InsERT.Moria.Kadry.Duze.WyplacanieWynagrodzenia.CoMiesiac;
                            if (baseUmowa.WyplacaWynagrodzenie != targetWyplata)
                            {
                                baseUmowa.WyplacaWynagrodzenie = targetWyplata;
                            }

                            if (baseUmowa.DataRachunkuPrzesuniecie != 25)
                            {
                                baseUmowa.DataRachunkuPrzesuniecie = 25;
                            }

                            if (baseUmowa.PrzesuniecieNaliczenia != 0)
                            {
                                baseUmowa.PrzesuniecieNaliczenia = 0;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogNexo.Informacja($"Ignorowano błąd harmonogramu wypłat (np. ograniczenia edycji typu umowy): {ex.Message}");
                        }

                        // Set rates (0.00m)
                        foreach (var sp in baseUmowa.SkladnikiPlacowe)
                        {
                            foreach (var val in sp.Wartosci)
                            {
                                val.Wartosc = 0.00m;
                            }
                        }

                        // Custom Flag (OPODATKOWANE / NIEOPODATKOWANE)
                        try
                        {
                            var flagName = isStudentUnder26 ? "NIEOPODATKOWANE" : "OPODATKOWANE";
                            var flagiWlasneMgr = uchwyt.PodajObiektTypu<InsERT.Moria.Flagi.IFlagiWlasne>();
                            var flag = flagiWlasneMgr.Dane.Wszystkie().FirstOrDefault(f => f.Nazwa == flagName);
                            if (flag != null)
                            {
                                baseUmowa.FlagaWlasna = flag;
                            }
                            else
                            {
                                item.Ostrzezenia += $"Nie znaleziono flagi własnej '{flagName}' w bazie. ";
                            }
                        }
                        catch (Exception ex)
                        {
                            item.Ostrzezenia += $"Błąd przypisywania flagi własnej: {ex.Message}. ";
                        }

                        if (!umowa.Zapisz())
                        {
                            var bledy = uchwyt.PodajBledy(umowa);
                            var msg = string.Join("; ", bledy.Select(x => $"[{x.Waznosc}] {x.Tresc}"));
                            LogNexo.Informacja($"Błąd zapisu umowy dla {item.Imie} {item.Nazwisko}: {msg}");
                            return WynikOperacji.ZakonczonaNiepowodzeniem($"Pracownik został dodany, ale błąd zapisu umowy: {msg}");
                        }

                        // ZUS Registration Declaration Automation (ZUA for Civil Contracts)
                        if (baseUmowa.RaportZUS == (byte)InsERT.Moria.Kadry.TypRaportuZUS.RCA)
                        {
                            try
                            {
                                var dekZusMgr = uchwyt.PodajObiektTypu<IDeklaracjeZUS>();
                                
                                // Query for an existing unexported zgłoszeniowa declaration (the last one)
                                var existingDecl = dekZusMgr.Dane.Wszystkie()
                                    .Where(d => d.RodzajDeklaracji == (byte)RodzajDeklaracjiZUS.Zgloszeniowa 
                                             && d.StatusWysylki == (byte)StatusWysylkiDeklaracjiZUS.Brak)
                                    .OrderByDescending(d => d.Id)
                                    .FirstOrDefault();

                                IDeklaracjaZUS deklBO;
                                if (existingDecl != null)
                                {
                                    deklBO = dekZusMgr.Znajdz(existingDecl);
                                }
                                else
                                {
                                    deklBO = dekZusMgr.UtworzDeklaracjeZgloszeniowa();
                                }

                                using (deklBO)
                                {
                                    bool alreadyAdded = deklBO.Dane.DeklaracjeZgloszeniowe
                                        .Any(z => z.UmowaPracownicza != null && z.UmowaPracownicza.Id == baseUmowa.Id);

                                    if (!alreadyAdded)
                                    {
                                        var zgloszenie = new DeklaracjaZgloszeniowa();
                                        deklBO.Dane.DeklaracjeZgloszeniowe.Add(zgloszenie);

                                        // Load fresh unattached entities from read-only repositories to avoid context mismatches
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

                                            if (!deklBO.Zapisz())
                                            {
                                                var bledy = uchwyt.PodajBledy(deklBO);
                                                var msg = string.Join("; ", bledy.Select(x => $"[{x.Waznosc}] {x.Tresc}"));
                                                item.Ostrzezenia += $"Nie udało się zapisać zgłoszenia ZUS: {msg}. ";
                                            }
                                            else
                                            {
                                                LogNexo.Informacja($"Dodano umowę pracownika {item.Imie} {item.Nazwisko} do zgłoszenia ZUS (Zgloszeniowa).");
                                            }
                                        }
                                        else
                                        {
                                            item.Ostrzezenia += "Nie udało się załadować powiązanych danych umowy/pracownika dla ZUS. ";
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                item.Ostrzezenia += $"Błąd podczas tworzenia/aktualizacji zgłoszenia ZUS: {ex.Message}. ";
                                LogNexo.Blad(ex);
                            }
                        }
                    }
                }

                if (!string.IsNullOrEmpty(item.Ostrzezenia))
                {
                    LogNexo.Informacja($"Zaimportowano pracownika {item.Imie} {item.Nazwisko} z ostrzeżeniami: {item.Ostrzezenia}");
                    return WynikOperacji.ZakonczonaPowodzeniem($"Zaimportowano z ostrzeżeniami: {item.Ostrzezenia}");
                }
                LogNexo.Informacja($"Zakończono pomyślnie import pracownika: {item.Imie} {item.Nazwisko}");
                return WynikOperacji.ZakonczonaPowodzeniem();
            }
            catch (ArgumentException argException)
            {
                LogNexo.Blad(argException);
                return WynikOperacji.ZakonczonaNiepowodzeniem($"Wyjątek argumentu podczas importu: {argException.Message}");
            }
            catch (Exception ex)
            {
                LogNexo.Blad(ex);
                return WynikOperacji.ZakonczonaNiepowodzeniem($"Wyjątek podczas importu: {ex.Message}");
            }
        }

        private Plec GetPlec(string plecStr)
        {
            if (string.IsNullOrWhiteSpace(plecStr)) return Plec.Nieokreslona;
            string lower = plecStr.ToLower();
            if (lower.Contains("kob") || lower.Contains("fem")) return Plec.Kobieta;
            if (lower.Contains("męż") || lower.Contains("mez") || lower.Contains("mal")) return Plec.Mezczyzna;
            return Plec.Nieokreslona;
        }

        private Panstwo ZnajdzPanstwo(IUchwyt uchwyt, string name)
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

        private Wojewodztwo? ZnajdzWojewodztwo(IUchwyt uchwyt, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            string norm = name.ToUpper();
            return uchwyt.Wojewodztwa().Dane.Wszystkie().FirstOrDefault(w => w.Nazwa.ToUpper() == norm);
        }

        private Powiat? ZnajdzPowiat(IUchwyt uchwyt, Wojewodztwo? woj, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            string norm = name.ToUpper();
            if (woj != null)
            {
                return uchwyt.Powiaty().Dane.Wszystkie().FirstOrDefault(p => p.Wojewodztwo.Id == woj.Id && p.Nazwa.ToUpper() == norm);
            }
            return uchwyt.Powiaty().Dane.Wszystkie().FirstOrDefault(p => p.Nazwa.ToUpper() == norm);
        }

        private Gmina? ZnajdzGmine(IUchwyt uchwyt, Powiat? pow, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            string norm = name.ToUpper();
            if (pow != null)
            {
                return uchwyt.Gminy().Dane.Wszystkie().FirstOrDefault(g => g.Powiat.Id == pow.Id && g.Nazwa.ToUpper() == norm);
            }
            return uchwyt.Gminy().Dane.Wszystkie().FirstOrDefault(g => g.Nazwa.ToUpper() == norm);
        }

        private OddzialNFZ? MappedNFZ(string nfzName)
        {
            if (string.IsNullOrWhiteSpace(nfzName)) return null;
            string lower = nfzName.ToLower();
            if (lower.Contains("lubel") || lower.Contains("lublin")) return OddzialNFZ.lublin;
            if (lower.Contains("lubus") || lower.Contains("zielon")) return OddzialNFZ.zielona;
            if (lower.Contains("łódź") || lower.Contains("lodz") || lower.Contains("łódz")) return OddzialNFZ.lodz;
            if (lower.Contains("małop") || lower.Contains("kraków") || lower.Contains("krakow")) return OddzialNFZ.krakow;
            if (lower.Contains("mazow") || lower.Contains("warszaw")) return OddzialNFZ.warszawa;
            if (lower.Contains("wielk") || lower.Contains("pozna")) return OddzialNFZ.poznan;
            if (lower.Contains("opol")) return OddzialNFZ.opole;
            if (lower.Contains("podkar") || lower.Contains("rzeszów") || lower.Contains("rzeszow")) return OddzialNFZ.rzeszow;
            if (lower.Contains("podla") || lower.Contains("białys") || lower.Contains("bialys")) return OddzialNFZ.bialystok;
            if (lower.Contains("pomor") || lower.Contains("gdań") || lower.Contains("gdan")) return OddzialNFZ.gdansk;
            if (lower.Contains("śląs") || lower.Contains("slas") || lower.Contains("katow")) return OddzialNFZ.katowice;
            if (lower.Contains("święt") || lower.Contains("swiet") || lower.Contains("kielc")) return OddzialNFZ.kielce;
            if (lower.Contains("warmi") || lower.Contains("olszt")) return OddzialNFZ.olsztyn;
            if (lower.Contains("zacho") || lower.Contains("szcze")) return OddzialNFZ.szczecin;
            if (lower.Contains("dolno") || lower.Contains("wrocł") || lower.Contains("wrocl")) 
                return (OddzialNFZ)Enum.Parse(typeof(OddzialNFZ), "wroclaw");
            if (lower.Contains("kujaw") || lower.Contains("bydgo")) 
                return (OddzialNFZ)Enum.Parse(typeof(OddzialNFZ), "bydgoszcz");
            return null;
        }

        internal static Panstwo? ResolvePanstwo(IUchwyt uchwyt, string inputName, out string platnikCitizenshipName)
        {
            platnikCitizenshipName = "";
            if (string.IsNullOrWhiteSpace(inputName)) return null;

            string normInput = inputName.Trim().ToLower();

            // 1. Search in our static CountryRegistry
            var countryInfo = CountryRegistry.FirstOrDefault(c => 
                c.PolishName.Equals(normInput, StringComparison.OrdinalIgnoreCase) || 
                c.CitizenshipAdjective.Equals(normInput, StringComparison.OrdinalIgnoreCase) ||
                c.IsoCode.Equals(normInput, StringComparison.OrdinalIgnoreCase)
            );

            string searchName = normInput;
            string isoCode = "";

            if (countryInfo != null)
            {
                searchName = countryInfo.PolishName;
                isoCode = countryInfo.IsoCode;
                platnikCitizenshipName = countryInfo.CitizenshipAdjective;
            }
            else
            {
                // Fallback guessing
                platnikCitizenshipName = normInput; // fallback
            }

            var panstwaMgr = uchwyt.PodajObiektTypu<InsERT.Moria.Klienci.IPanstwa>();
            
            // 2. Search database by name
            var dbPanstwo = panstwaMgr.Dane.Wszystkie().FirstOrDefault(p => 
                p.Nazwa.Equals(searchName, StringComparison.OrdinalIgnoreCase) || 
                p.KodPanstwaUE.Equals(searchName, StringComparison.OrdinalIgnoreCase)
            );

            if (dbPanstwo != null)
            {
                return dbPanstwo;
            }

            // 3. If country is not in database, create it automatically
            if (!string.IsNullOrEmpty(isoCode))
            {
                try
                {
                    using (var panstwoBO = panstwaMgr.Utworz())
                    {
                        panstwoBO.Dane.Nazwa = searchName;
                        panstwoBO.Dane.KodPanstwaUE = isoCode;
                        panstwoBO.Dane.CzlonekUE = false;

                        if (panstwoBO.Zapisz())
                        {
                            LogNexo.Informacja($"Utworzono brakujące państwo: {searchName} (ISO: {isoCode})");
                            return panstwoBO.Dane;
                        }
                        else
                        {
                            var bledy = uchwyt.PodajBledy(panstwoBO);
                            var msg = string.Join("; ", bledy.Select(x => x.Tresc));
                            LogNexo.Informacja($"Błąd zapisu państwa {searchName}: {msg}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogNexo.Blad(ex);
                }
            }

            // Ultimate fallback
            return panstwaMgr.Dane.Wszystkie().FirstOrDefault(p => p.KodPanstwaUE == "PL");
        }

        internal static Obywatelstwo? ResolveObywatelstwo(IUchwyt uchwyt, string inputName)
        {
            if (string.IsNullOrWhiteSpace(inputName)) return null;

            string platnikName;
            var panstwo = ResolvePanstwo(uchwyt, inputName, out platnikName);
            if (panstwo == null) return null;

            if (string.IsNullOrEmpty(platnikName))
            {
                platnikName = panstwo.Nazwa.ToLower() + "skie"; // fallback guess
            }

            var obywatelstwaMgr = uchwyt.PodajObiektTypu<InsERT.Moria.Klienci.IObywatelstwa>();
            
            // 1. Search existing citizenship in DB
            var dbObywatelstwo = obywatelstwaMgr.Dane.Wszystkie().FirstOrDefault(o => 
                o.Nazwa.Equals(platnikName, StringComparison.OrdinalIgnoreCase) || 
                (o.Panstwo != null && o.Panstwo.Id == panstwo.Id)
            );

            if (dbObywatelstwo != null)
            {
                return dbObywatelstwo;
            }

            // 2. If not found, create new citizenship in DB
            try
            {
                using (var obyBO = obywatelstwaMgr.Utworz())
                {
                    obyBO.Dane.Nazwa = platnikName;
                    obyBO.Dane.Panstwo = panstwo;

                    if (obyBO.Zapisz())
                    {
                        LogNexo.Informacja($"Utworzono brakujące obywatelstwo: {platnikName}");
                        return obyBO.Dane;
                    }
                    else
                    {
                        var bledy = uchwyt.PodajBledy(obyBO);
                        var msg = string.Join("; ", bledy.Select(x => x.Tresc));
                        LogNexo.Informacja($"Błąd zapisu obywatelstwa {platnikName}: {msg}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogNexo.Blad(ex);
            }

            return null;
        }

        private static string CleanAndStemUs(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            
            var stopWords = new HashSet<string> { "urzad", "skarbowy", "w", "we", "us", "podatkowy", "izba", "dla" };
            
            var words = input.ToLower().Split(new[] { ' ', '-', '/' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(w => w.Replace("ą", "a").Replace("ć", "c").Replace("ę", "e")
                              .Replace("ł", "l").Replace("ń", "n").Replace("ó", "o")
                              .Replace("ś", "s").Replace("ź", "z").Replace("ż", "z"))
                .Where(w => !stopWords.Contains(w))
                .Select(w => {
                    w = w.Replace("pierwszy", "1").Replace("drugi", "2").Replace("trzeci", "3");
                    return w.Length > 4 ? w.Substring(0, 4) : w;
                });

            return string.Join("", words);
        }

        internal static OrganPodatkowy? ResolveOrganPodatkowy(IUchwyt uchwyt, string usName)
        {
            if (string.IsNullOrWhiteSpace(usName)) return null;

            string stemName = CleanAndStemUs(usName);
            if (string.IsNullOrEmpty(stemName)) return null;

            // 1. Search existing OrganPodatkowy in the customer database
            var localUS = uchwyt.Podmioty().Dane.WszystkieOrganyPodatkowe()
                .AsEnumerable()
                .FirstOrDefault(o => o.Firma != null && 
                                     (CleanAndStemUs(o.Firma.Nazwa) == stemName || 
                                      CleanAndStemUs(o.NazwaSkrocona) == stemName));

            if (localUS != null)
            {
                return localUS.Firma.OrganPodatkowy;
            }

            // 2. Search built-in template database using DbContext
            try
            {
                var podmiotyMgr = uchwyt.Podmioty();
                
                string kod = null;
                string templateNazwa = null;
                
                using (var bo = podmiotyMgr.UtworzPracownika())
                {
                    var getUow = (InsERT.Mox.BusinessObjects.IGetUnitOfWork)bo;
                    var scope = (InsERT.Mox.Runtime.IInjectionScope)getUow.UnitOfWork;
                    var context = (ModelDanychContainer)scope.ScopedContainer.GetObject(typeof(ModelDanychContainer));
                    
                    var templateUS = context.UrzedySkarbowe
                        .AsEnumerable()
                        .FirstOrDefault(t => t.Nazwa != null && CleanAndStemUs(t.Nazwa) == stemName);

                    if (templateUS != null)
                    {
                        kod = templateUS.Kod;
                        templateNazwa = templateUS.Nazwa;
                    }
                }

                if (kod != null)
                {
                    // 3. Create a new OrganPodatkowy based on the 4-digit code
                    using (var organBO = podmiotyMgr.UtworzOrganPodatkowy())
                    {
                        var organ = (IOrganPodatkowy)organBO;
                        if (organ.WypelnijDlaKodu(kod))
                        {
                            organBO.AutoSymbol();
                            if (organBO.Zapisz())
                            {
                                LogNexo.Informacja($"Utworzono brakujący urząd skarbowy: {templateNazwa} (Kod: {kod})");
                                return organBO.Dane.Firma.OrganPodatkowy;
                            }
                            else
                            {
                                var bledy = uchwyt.PodajBledy(organBO);
                                var msg = string.Join("; ", bledy.Select(x => x.Tresc));
                                LogNexo.Informacja($"Błąd zapisu urzędu skarbowego {templateNazwa}: {msg}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ResolveOrganPodatkowy ERROR]: {ex}");
                LogNexo.Blad(ex);
            }

            return null;
        }

        private class CountryInfo
        {
            public string PolishName { get; }
            public string IsoCode { get; }
            public string CitizenshipAdjective { get; }

            public CountryInfo(string polishName, string isoCode, string citizenshipAdjective)
            {
                PolishName = polishName;
                IsoCode = isoCode;
                CitizenshipAdjective = citizenshipAdjective;
            }
        }

        private static readonly List<CountryInfo> CountryRegistry = new List<CountryInfo>
        {
            new CountryInfo("Polska", "PL", "polskie"),
            new CountryInfo("Ukraina", "UA", "ukraińskie"),
            new CountryInfo("Białoruś", "BY", "białoruskie"),
            new CountryInfo("Gruzja", "GE", "gruzińskie"),
            new CountryInfo("Mołdawia", "MD", "mołdawskie"),
            new CountryInfo("Rosja", "RU", "rosyjskie"),
            new CountryInfo("Paragwaj", "PY", "paragwajskie"),
            new CountryInfo("Zimbabwe", "ZW", "zimbabweńskie"),
            new CountryInfo("Indie", "IN", "indyjskie"),
            new CountryInfo("Wietnam", "VN", "wietnamskie"),
            new CountryInfo("Kazachstan", "KZ", "kazachskie"),
            new CountryInfo("Uzbekistan", "UZ", "uzbekistańskie"),
            new CountryInfo("Armenia", "AM", "armeńskie"),
            new CountryInfo("Filipiny", "PH", "filipińskie"),
            new CountryInfo("Niemcy", "DE", "niemieckie"),
            new CountryInfo("Czechy", "CZ", "czeskie"),
            new CountryInfo("Słowacja", "SK", "słowackie"),
            new CountryInfo("Turcja", "TR", "tureckie")
        };
    }
}
