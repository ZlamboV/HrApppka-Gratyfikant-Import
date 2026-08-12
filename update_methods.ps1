$filePath = "c:\ NEXO SDK SFERA\HrAppka Import Pracowników\OperacjaImportuPracownikow20.cs"
$content = Get-Content -Raw $filePath

$preAnalyzeRegex = "(?s)private void PreAnalyzeRows\(IUchwyt uchwyt, List<PracownikImportRow> rows\).*?private WynikOperacji ZaimportujRow20"

$newMethods = @"
private void PreAnalyzeRows(IUchwyt uchwyt, List<PracownikImportRow> rows)
        {
            var podmioty = uchwyt.Podmioty();
            var existingPesels = podmioty.Dane.WszyscyPracownicy()
                .Select(p => p.Osoba.PESEL)
                .Where(p => !string.IsNullOrEmpty(p))
                .ToHashSet();

            var existingPassports = podmioty.Dane.WszyscyPracownicy()
                .Select(p => p.Osoba.SeriaINumerDokumentuTozsamosci)
                .Where(p => !string.IsNullOrEmpty(p))
                .ToHashSet();

            var urzedyMgr = uchwyt.PodajObiektTypu<InsERT.Moria.Instytucje.IInstytucje>();
            var wszystkieUrzedy = urzedyMgr.Dane.Wszystkie().Where(i => i.JestOrganemPodatkowym).Select(i => i.OrganPodatkowy).ToList();

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
                            var existingOsoba = podmioty.Dane.WszystkieOsoby()
                                .FirstOrDefault(p => p.Osoba != null && p.Osoba.SeriaINumerDokumentuTozsamosci == row.Paszport);
                            SprawdzIstniejacyPracownikUmowy(uchwyt, existingOsoba, row, "Paszport", row.Paszport);
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
                        row.Status = "B³êdny PESEL";
                    }
                    else
                    {
                        if (!string.IsNullOrWhiteSpace(row.BirthDateExcel) && DateTime.TryParse(row.BirthDateExcel, out DateTime dtExcel))
                        {
                            if (dtExcel.Date != birthDate.Value.Date)
                            {
                                var chosenDate = WpfWarningResolverDialog.RozwiazDataUrodzenia(row.Nazwisko, row.Imie, birthDate.Value, dtExcel);
                                row.ResolvedBirthDate = chosenDate ?? birthDate.Value;
                            }
                        }
                        if (!string.IsNullOrWhiteSpace(row.PlecExcel))
                        {
                            if (!string.Equals(row.PlecExcel, gender, StringComparison.OrdinalIgnoreCase))
                            {
                                var chosenGender = WpfWarningResolverDialog.RozwiazPlec(row.Nazwisko, row.Imie, gender, row.PlecExcel);
                                row.ResolvedGender = chosenGender ?? gender;
                            }
                        }
                    }

                    if (existingPesels.Contains(row.Pesel))
                    {
                        var existingOsoba = podmioty.Dane.WszystkieOsoby()
                            .FirstOrDefault(p => p.Osoba != null && p.Osoba.PESEL == row.Pesel);
                        SprawdzIstniejacyPracownikUmowy(uchwyt, existingOsoba, row, "PESEL", row.Pesel);
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

                if (!string.IsNullOrWhiteSpace(row.UrzadSkarbowy))
                {
                    var matchingUs = wszystkieUrzedy.FirstOrDefault(u => u.Nazwa != null && u.Nazwa.Contains(row.UrzadSkarbowy, StringComparison.OrdinalIgnoreCase));
                    if (matchingUs == null)
                    {
                        row.SelectedOrganPodatkowy = WpfWarningResolverDialog.RozwiazBrakUrzeduSkarbowego(row.Nazwisko, row.Imie, row.UrzadSkarbowy, wszystkieUrzedy);
                    }
                    else
                    {
                        row.SelectedOrganPodatkowy = matchingUs;
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
                row.Status = "Pominiêty (Istnieje)";
                row.Ostrzezenia += "Wiersz zosta³ pominiêty na ¿¹danie u¿ytkownika. ";
            }
            else if (decyzja == DecyzjaIstniejacyPracownik.NadpiszUmowe)
            {
                row.Status = "Aktualizacja (Nadpisanie)";
                row.Ostrzezenia += "U¿ytkownik zdecydowa³ o nadpisaniu umowy. ";
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

        private WynikOperacji ZaimportujRow20"

$content = $content -replace $preAnalyzeRegex, $newMethods
Set-Content $filePath $content
