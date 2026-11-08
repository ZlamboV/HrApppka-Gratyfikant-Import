using System;
using System.Collections.Generic;
using System.Linq;
using InsERT.Moria.DeklaracjeZUS;
using InsERT.Moria.Kadry.Duze;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Sfera;

namespace HrAppka_Import_Pracowników.Services
{
    public record ZusRegistrationRequest(
        UmowaPracownicza Contract,
        Pracownik Employee,
        bool IsZzaOnly,
        DateTime StartDate,
        bool PlacEmerytalna,
        bool PlacRentowa,
        bool PlacChorobowa,
        bool PlacWypadkowa
    );

    public class ZusBatchResult
    {
        public bool Success { get; set; }
        public int TotalProcessed { get; set; }
        public int SuccessfullyAdded { get; set; }
        public int Skipped { get; set; }
        public List<string> Warnings { get; } = new();
        public List<string> Errors { get; } = new();
        public int? DeclarationId { get; set; }
    }

    /// <summary>
    /// Serwis zbiorczego generowania i pojedynczego zapisu deklaracji zgłoszeniowych ZUS (ZUA / ZZA).
    /// Zgodny z Wymaganiem R7 (ORIGINAL_REQUEST.md §R7, PROJECT.md §Feature 19).
    /// </summary>
    public static class ZusDeclarationBatchService
    {
        public static ZusBatchResult ProcessAndSaveBatch(
            IUchwyt uchwyt, 
            IEnumerable<ZusRegistrationRequest> requests,
            Action<string>? logger = null)
        {
            var result = new ZusBatchResult();
            var requestList = requests?.ToList() ?? new List<ZusRegistrationRequest>();
            result.TotalProcessed = requestList.Count;

            if (requestList.Count == 0)
            {
                result.Success = true;
                return result;
            }

            if (uchwyt == null) throw new ArgumentNullException(nameof(uchwyt));

            try
            {
                var dekZusMgr = uchwyt.DeklaracjeZUS();

                // 1. Zlokalizuj istniejący, niewysłany projekt deklaracji zgłoszeniowej lub utwórz nowy
                var existingDecl = dekZusMgr.Dane.Wszystkie()
                    .Where(d => d.RodzajDeklaracji == (byte)RodzajDeklaracjiZUS.Zgloszeniowa 
                             && d.StatusWysylki == (byte)StatusWysylkiDeklaracjiZUS.Brak
                             && !d.WyslanyDoPlatnika)
                    .OrderByDescending(d => d.Id)
                    .FirstOrDefault();

                using (IDeklaracjaZUS deklBO = existingDecl != null 
                    ? dekZusMgr.Znajdz(existingDecl) 
                    : dekZusMgr.UtworzDeklaracjeZgloszeniowa())
                {
                    if (deklBO is not DeklaracjaZUSBO concreteBO)
                    {
                        result.Errors.Add("Obiekt biznesowy deklaracji ZUS nie implementuje typu DeklaracjaZUSBO.");
                        result.Success = false;
                        return result;
                    }

                    var podmioty = uchwyt.Podmioty();
                    var umowyMgr = uchwyt.UmowyPracowniczeGr();

                    foreach (var req in requestList)
                    {
                        try
                        {
                            if (req.Contract == null || req.Employee == null)
                            {
                                result.Skipped++;
                                result.Warnings.Add("Żądanie zgłoszenia ZUS zawiera puste odwołanie do umowy lub pracownika.");
                                continue;
                            }

                            // Unikaj duplikowania tego samego zgłoszenia w tej samej deklaracji
                            bool alreadyPresent = deklBO.Dane.DeklaracjeZgloszeniowe
                                .Any(z => z.UmowaPracownicza != null && z.UmowaPracownicza.Id == req.Contract.Id);

                            if (alreadyPresent)
                            {
                                result.Skipped++;
                                result.Warnings.Add($"Umowa ID {req.Contract.Id} znajduje się już w bieżącym zgłoszeniu ZUS.");
                                continue;
                            }

                            // Pobierz świeże encje powiązane z bieżącym kontekstem Unit of Work
                            var freshEmployeePodmiot = podmioty.Dane.WszystkieOsoby().FirstOrDefault(p => p.Id == req.Employee.Id);
                            var freshContract = umowyMgr.Dane.Wszystkie().FirstOrDefault(u => u.Id == req.Contract.Id);
                            var freshEmployee = freshEmployeePodmiot?.Osoba?.Pracownik;

                            if (freshContract == null || freshEmployee == null)
                            {
                                result.Skipped++;
                                result.Warnings.Add($"Nie odnaleziono powiązania pracownik-umowa w bazie dla umowy ID {req.Contract.Id}. Pominięto zgłoszenie ZUS.");
                                continue;
                            }

                            var osoba = freshEmployee.Osoba ?? freshEmployeePodmiot?.Osoba;
                            bool hasPesel = !string.IsNullOrWhiteSpace(osoba?.PESEL);
                            bool hasDoc = osoba?.TypDokumentuTozsamosci.HasValue == true &&
                                          !string.IsNullOrWhiteSpace(osoba?.SeriaINumerDokumentuTozsamosci);

                            if (!hasPesel && !hasDoc)
                            {
                                result.Skipped++;
                                result.Warnings.Add($"Pracownik {osoba?.Imie} {osoba?.Nazwisko} nie posiada ani numeru PESEL, ani dokumentu tożsamości. Pominięto zgłoszenie ZUS.");
                                continue;
                            }

                            var zgloszenie = new DeklaracjaZgloszeniowa
                            {
                                UmowaPracownicza = freshContract,
                                Pracownik = freshEmployee,
                                ImieNazwisko = $"{osoba?.Imie} {osoba?.Nazwisko}".ToUpper().Trim()
                            };

                            DateTime obligationDate = freshContract.OkresObowiazywania?.DataPoczatkowa ?? req.StartDate;

                            if (req.IsZzaOnly)
                            {
                                zgloszenie.Typ = (byte)TypDeklaracjiZgloszeniowej.ZZA;
                                var daneZza = new DaneZZA
                                {
                                    TYPZgloszenia = 0,
                                    DataObowiazkuUbezpieczenZdrowotnych = obligationDate
                                };
                                concreteBO.WypelnijDeklaracjeZgloszeniona(zgloszenie, daneZza);
                            }
                            else
                            {
                                zgloszenie.Typ = (byte)TypDeklaracjiZgloszeniowej.ZUA;
                                var daneZua = new DaneZUA
                                {
                                    TYPZgloszenia = 0,
                                    DataObowiazkuUbezpieczenSpoleczych = obligationDate,
                                    DataObowiazkuUbezpieczenZdrowotnych = obligationDate,
                                    PlacSkladkaEmerytalna = req.PlacEmerytalna,
                                    PlacSkladkaRentowa = req.PlacRentowa,
                                    PlacSkladkaChorobowa = req.PlacChorobowa,
                                    PlacSkladkaWypadkowa = req.PlacWypadkowa
                                };
                                concreteBO.WypelnijDeklaracjeZgloszeniona(zgloszenie, daneZua);
                            }

                            deklBO.Dane.DeklaracjeZgloszeniowe.Add(zgloszenie);
                            result.SuccessfullyAdded++;
                        }
                        catch (Exception itemEx)
                        {
                            result.Skipped++;
                            result.Warnings.Add($"Błąd przy przygotowaniu zgłoszenia ZUS dla umowy ID {req.Contract.Id}: {itemEx.Message}");
                        }
                    }

                    // 2. Wykonaj pojedynczy zbiorczy zapis (Single-Save Batch) jeśli dodano zgłoszenia
                    if (result.SuccessfullyAdded > 0)
                    {
                        if (!deklBO.Zapisz())
                        {
                            var bledy = uchwyt.PodajBledy(deklBO);
                            var errorDetails = bledy != null && bledy.Any()
                                ? string.Join("; ", bledy.Select(b => $"[{b.Waznosc}] {b.Tresc}"))
                                : string.Join("; ", deklBO.PobierzKomunikatyBledow().Select(b => b.Tresc));

                            result.Errors.Add($"Nie udało się zapisać zbiorczej deklaracji ZUS: {errorDetails}");
                            result.Success = false;
                        }
                        else
                        {
                            result.DeclarationId = deklBO.Dane.Id;
                            result.Success = true;
                            logger?.Invoke($"Zapisano zbiorczą deklarację ZUS (ID: {deklBO.Dane.Id}) zawierającą {result.SuccessfullyAdded} zgłoszeń.");
                        }
                    }
                    else
                    {
                        result.Success = true;
                    }
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add($"Błąd krytyczny podczas zbiorczego zapisu deklaracji ZUS: {ex.Message}");
                result.Success = false;
            }

            return result;
        }
    }
}
