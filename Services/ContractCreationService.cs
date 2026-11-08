using System;

using System.Collections.Generic;

using System.Globalization;

using System.Linq;

using InsERT.Moria.Flagi;

using InsERT.Moria.Kadry;

using InsERT.Moria.Kadry.Duze;

using InsERT.Moria.ModelDanych;

using InsERT.Moria.Rozszerzanie;

using InsERT.Moria.Sfera;



namespace HrAppka_Import_Pracowników.Services

{

    public class SzablonUmowyItem

    {

        public int Id { get; set; }

        public string Nazwa { get; set; } = string.Empty;

        public string Opis { get; set; } = string.Empty;

        public bool IsDefault { get; set; }

        public int TypUmowyId { get; set; }

        public string TypUmowyNazwa { get; set; } = string.Empty;

        public bool IsCivil { get; set; }



        public override string ToString() => Nazwa;

    }



    public interface IContractCreationService

    {

        List<SzablonUmowyItem> GetAvailableTemplates(IUchwyt uchwyt);

        SzablonUmowyItem? GetDefaultTemplate(IUchwyt uchwyt, string? factoryOrDepartmentName = null);

        UmowaPracowniczaGr CreateContract(

            IUchwyt uchwyt,

            Podmiot employeePodmiot,

            PracownikImportRow row,

            int templateId,

            JednostkaOrganizacyjnaGr? department = null,

            decimal? wageRate = null);

    }



    /// <summary>

    /// Serwis zarządzania szablonami umów i instancjonowania umów pracowniczych / cywilnoprawnych w Sfera SDK 62.

    /// Realizuje Wymaganie R2 (ISzablonyUmowPracowniczych, IUmowyPracownicze.Utworz(szablon)).

    /// </summary>

    public class ContractCreationService : IContractCreationService

    {

        private static readonly Lazy<ContractCreationService> _lazy =

            new Lazy<ContractCreationService>(() => new ContractCreationService());



        public static ContractCreationService Instance => _lazy.Value;



        public List<SzablonUmowyItem> GetAvailableTemplates(IUchwyt uchwyt)

        {

            if (uchwyt == null) throw new ArgumentNullException(nameof(uchwyt));



            var szablonyMgr = uchwyt.SzablonyUmowPracowniczychGr();

            return szablonyMgr.Dane.Wszystkie()

                .OrderBy(s => s.Nazwa)

                .AsEnumerable()

                .Where(s => !s.IsInRecycleBin)

                .Select(s => new SzablonUmowyItem

                {

                    Id = s.Id,

                    Nazwa = s.Nazwa,

                    Opis = s.Opis ?? string.Empty,

                    IsDefault = s.DomyslnyWTypie,

                    TypUmowyId = s.TypUmowyPracowniczejGr != null ? s.TypUmowyPracowniczejGr.Id : 0,

                    TypUmowyNazwa = s.TypUmowyPracowniczejGr != null ? s.TypUmowyPracowniczejGr.Nazwa : string.Empty,

                    IsCivil = s.TypUmowyPracowniczejGr != null && s.TypUmowyPracowniczejGr.Typ == (byte)TypUmowyPracowniczej.Cywilnoprawna

                })

                .ToList();

        }



        public SzablonUmowyItem? GetDefaultTemplate(IUchwyt uchwyt, string? factoryOrDepartmentName = null)

        {

            var templates = GetAvailableTemplates(uchwyt);



            // If factory/department name is provided, search if a dedicated template exists

            if (!string.IsNullOrWhiteSpace(factoryOrDepartmentName))

            {

                var factoryMatch = templates.FirstOrDefault(t => t.Nazwa.IndexOf(factoryOrDepartmentName.Trim(), StringComparison.OrdinalIgnoreCase) >= 0);

                if (factoryMatch != null) return factoryMatch;

            }



            // Prioritize civil contract / Umowa zlecenie for HrAppka import:

            return templates.FirstOrDefault(t => t.IsCivil && string.Equals(t.Nazwa.Trim(), "Umowa zlecenie", StringComparison.OrdinalIgnoreCase))

                ?? templates.FirstOrDefault(t => t.Nazwa.IndexOf("zlecenie", StringComparison.OrdinalIgnoreCase) >= 0)

                ?? templates.FirstOrDefault(t => t.IsCivil && t.IsDefault)

                ?? templates.FirstOrDefault(t => t.IsCivil)

                ?? templates.FirstOrDefault(t => t.IsDefault)

                ?? templates.FirstOrDefault();

        }



        public UmowaPracowniczaGr CreateContract(

            IUchwyt uchwyt,

            Podmiot employeePodmiot,

            PracownikImportRow row,

            int templateId,

            JednostkaOrganizacyjnaGr? department = null,

            decimal? wageRate = null)

        {

            if (uchwyt == null) throw new ArgumentNullException(nameof(uchwyt));

            if (employeePodmiot == null) throw new ArgumentNullException(nameof(employeePodmiot));

            if (row == null) throw new ArgumentNullException(nameof(row));



            var szablonyMgr = uchwyt.SzablonyUmowPracowniczychGr();



            // If templateId is <= 0, resolve default template

            if (templateId <= 0)

            {

                var defaultItem = GetDefaultTemplate(uchwyt);

                if (defaultItem != null)

                {

                    templateId = defaultItem.Id;

                }

            }



            var templateEntity = szablonyMgr.Dane.Wszystkie()

                .Where(s => s.Id == templateId)

                .AsEnumerable()

                .FirstOrDefault(s => !s.IsInRecycleBin);



            if (templateEntity == null)

            {

                throw new InvalidOperationException($"Nie odnaleziono aktywnego szablonu umowy o ID: {templateId}");

            }



            // Dates parsing and validation

            DateTime dtStart = DateTime.TryParse(row.DataRozpoczecia, out var ds) ? ds : DateTime.Today;

            DateTime? dtEnd = DateTime.TryParse(row.DataZakonczenia, out var de) ? de : null;

            DateTime dtSign = DateTime.TryParse(row.DataZawarcia, out var dw) ? dw : dtStart;



            if (dtEnd.HasValue && dtEnd.Value < dtStart)

            {

                throw new ArgumentException("Data zakończenia nie może być wcześniejsza niż data rozpoczęcia.");

            }



            var umowyMgr = uchwyt.UmowyPracowniczeGr();

            using (var umowaBO = umowyMgr.Utworz(templateEntity))

            {

                var u = umowaBO.Dane;



                // Resolve the Business Object's own EF DbContext (ModelDanychContainer).

                // Querying referenced entities (Pracownik, JednostkaOrganizacyjna, Kalendarz, etc.)

                // from this context guarantees they are attached to the exact ObjectContext controlling

                // data modification, eliminating cross-context detached entity exceptions on Zapisz().

                ModelDanychContainer? context = null;

                try

                {

                    var prop = u.GetType().GetProperty("ObjectContext", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);

                    context = prop?.GetValue(u) as ModelDanychContainer;

                }

                catch { }



                if (context == null)

                {

                    try

                    {

                        if (umowaBO is InsERT.Mox.BusinessObjects.IGetUnitOfWork getUow &&

                            getUow.UnitOfWork is InsERT.Mox.Runtime.IInjectionScope scope)

                        {

                            context = scope.ScopedContainer?.GetObject(typeof(ModelDanychContainer)) as ModelDanychContainer;

                        }

                    }

                    catch { }

                }



                var empId = employeePodmiot.Osoba?.Pracownik?.Id;

                if (context != null && empId.HasValue && empId.Value > 0)

                {

                    u.Pracownik = context.Pracownicy.FirstOrDefault(p => p.Id == empId.Value) 

                        ?? employeePodmiot.Osoba?.Pracownik;

                }

                else

                {

                    u.Pracownik = employeePodmiot.Osoba?.Pracownik;

                }



                // 1. Period of validity and signing date

                if (u.OkresObowiazywania == null)

                    u.OkresObowiazywania = new OkresWymaganyLewostronnie();

                u.OkresObowiazywania.DataPoczatkowa = dtStart;

                u.OkresObowiazywania.DataKoncowa = dtEnd;

                u.DataSporzadzenia = dtSign;



                // 2. Working Time Ratio & Calendar ("Bez cyklu", 1/1) - only for contract types that support calendars/dimensions

                bool supportsKalendarz = false;

                try

                {

                    if (u.TypUmowyPracowniczej != null)

                    {

                        supportsKalendarz = u.TypUmowyPracowniczej.Kalendarz;

                    }

                    else if (templateEntity?.TypUmowyPracowniczejGr != null)

                    {

                        supportsKalendarz = templateEntity.TypUmowyPracowniczejGr.Kalendarz;

                    }

                    else

                    {

                        // Mock fallback for unit test environments

                        supportsKalendarz = true;

                    }

                }

                catch

                {

                    supportsKalendarz = false;

                }



                if (supportsKalendarz)

                {

                    var kal = u.Kalendarze.FirstOrDefault();

                    if (kal == null)

                    {

                        var defaultKalendarz = (context != null 

                            ? (context.Kalendarze.FirstOrDefault(k => k.Nazwa == "Bez cyklu") ?? context.Kalendarze.FirstOrDefault())

                            : (uchwyt.Kalendarze().Dane.Aktywne().FirstOrDefault(k => k.Nazwa == "Bez cyklu") ?? uchwyt.Kalendarze().Dane.Aktywne().FirstOrDefault()));



                        kal = new KalendarzWUmowiePracowniczej

                        {

                            DataOd = dtStart,

                            Lp = 1,

                            Kalendarz = defaultKalendarz

                        };

                        u.Kalendarze.Add(kal);

                    }

                    else

                    {

                        kal.DataOd = dtStart;

                        if (kal.Kalendarz == null)

                        {

                            kal.Kalendarz = (context != null 

                                ? (context.Kalendarze.FirstOrDefault(k => k.Nazwa == "Bez cyklu") ?? context.Kalendarze.FirstOrDefault())

                                : (uchwyt.Kalendarze().Dane.Aktywne().FirstOrDefault(k => k.Nazwa == "Bez cyklu") ?? uchwyt.Kalendarze().Dane.Aktywne().FirstOrDefault()));

                        }

                    }



                    if (kal.WymiarZatrudnienia == null)

                        kal.WymiarZatrudnienia = new WymiarZatrudnienia { Licznik = 1, Mianownik = 1 };

                    else

                    {

                        kal.WymiarZatrudnienia.Licznik = 1;

                        kal.WymiarZatrudnienia.Mianownik = 1;

                    }

                }

                else

                {

                    u.Kalendarze.Clear();

                }



                // 3. Department & Job Position (Stanowiska)

                JednostkaOrganizacyjnaGr? attachedDept = null;

                if (department != null)

                {

                    if (context != null)

                    {

                        attachedDept = context.JednostkiOrganizacyjneGr.FirstOrDefault(d => d.Id == department.Id);

                    }

                    if (attachedDept == null)

                    {

                        try

                        {

                            attachedDept = uchwyt.JednostkiOrganizacyjne()?.Dane?.Wszystkie()?.FirstOrDefault(d => d.Id == department.Id) ?? department;

                        }

                        catch

                        {

                            attachedDept = department;

                        }

                    }

                }



                bool supportsStanowisko = false;

                bool supportsDzial = true;

                try

                {

                    var typUmowy = u.TypUmowyPracowniczej?.TypUmowyPracowniczejDlaKtoregoAktualny

                        ?? u.TypUmowyPracowniczej?.TypUmowyPracowniczejDlaKtoregoHistoria

                        ?? templateEntity?.TypUmowyPracowniczejGr;



                    if (typUmowy != null)

                    {

                        supportsStanowisko = typUmowy.Stanowisko;

                        supportsDzial = typUmowy.Dzial;

                    }

                    else

                    {

                        // Mock fallback for unit test environments

                        supportsStanowisko = true;

                        supportsDzial = true;

                    }

                }

                catch

                {

                    supportsStanowisko = false;

                    supportsDzial = true;

                }



                if (supportsDzial && attachedDept != null)

                {

                    var st = u.Stanowiska.FirstOrDefault();

                    if (st == null)

                    {

                        st = new StanowiskoWUmowiePracowniczej

                        {

                            DataOd = dtStart,

                            LP = 1

                        };

                        // Stanowisko entity must be attached to u.Stanowiska BEFORE assigning JednostkaOrganizacyjna,

                        // otherwise st is not sponsored by umowaBO's DataDomain and Mox throws UnsponsoredModificationException.

                        u.Stanowiska.Add(st);

                    }

                    else

                    {

                        st.DataOd = dtStart;

                    }



                    try

                    {

                        st.JednostkaOrganizacyjna = attachedDept;

                    }

                    catch (Exception exDept)

                    {

                        LogNexo.Informacja($"[Umowa] Nie udało się przypisać jednostki organizacyjnej do umowy: {exDept.Message}");

                    }

                }

                else if (!supportsStanowisko)

                {

                    u.Stanowiska.Clear();

                }



                // 4. Age, Student Status & Insurance Contributions

                DateTime birthDate = row.ResolvedBirthDate ?? (DateTime.TryParse(row.BirthDateExcel, out var bd) ? bd : DateTime.MinValue);

                int age = PeselHelper.CalculateAge(birthDate, dtStart);

                bool studentFlag = string.Equals(row.JestStudentem?.Trim(), "Tak", StringComparison.OrdinalIgnoreCase);

                bool ageUnder26 = age < 26;

                bool isStudentUnder26 = studentFlag && ageUnder26;

                bool isZusSubject = !isStudentUnder26;



                u.RaportZUS = (byte)(isStudentUnder26 ? TypRaportuZUS.Brak : TypRaportuZUS.RCA);

                u.NaliczaSkladkiEmerytalne = isZusSubject;

                u.NaliczaSkladkiRentowe = isZusSubject;

                u.NaliczaSkladkiZdrowotne = isZusSubject;

                u.NaliczaSkladkiWypadkowe = isZusSubject;

                u.NaliczaSkladkiFP = isZusSubject;

                u.NaliczaSkladkiFGSP = isZusSubject;

                u.NaliczaSkladkiChorobowe = false;



                u.StosujZwolnienieZPITDo26Lat = null;

                if (u.NaliczaSkladkiZdrowotne) u.NieOgraniczaZdrowotnego = false;

                if (u.NaliczaSkladkiChorobowe) u.NieOgraniczaChorobowego = false;

                if (u.NaliczaSkladkiFP) u.NaliczaSkladkiFPTylkoPowyzejMinimalnej = isZusSubject;



                // PIT Settings

                try

                {

                    var targetZaliczka = (byte)(ageUnder26 

                        ? SposobPobieraniaZaliczkiNaPodatekNaUmowie.WgUstawienPracownika 

                        : SposobPobieraniaZaliczkiNaPodatekNaUmowie.NiePobierajDoKwotyWolnej);

                    if (u.ZaliczkaNaPodatek != targetZaliczka) u.ZaliczkaNaPodatek = targetZaliczka;



                    var targetUlga = (byte)SposobOdliczaniaUlgiPodatkowejNaUmowie.WgUstawienPracownika;

                    if (u.UlgaPodatkowa != targetUlga) u.UlgaPodatkowa = targetUlga;

                }

                catch { }



                // Schedule: Monthly payment, pay date shifted 25 days

                try

                {

                    var targetWyplata = (byte)WyplacanieWynagrodzenia.CoMiesiac;

                    if (u.WyplacaWynagrodzenie != targetWyplata) u.WyplacaWynagrodzenie = targetWyplata;

                    if (u.DataRachunkuPrzesuniecie != 25) u.DataRachunkuPrzesuniecie = 25;

                    if (u.PrzesuniecieNaliczenia != 0) u.PrzesuniecieNaliczenia = 0;

                }

                catch { }



                // 5. Wage Rate (from wageRate param, row.Stawka via reflection if present, or 0.00m default for civil contracts)

                decimal parsedRate = 0.00m;

                bool hasRate = false;



                if (wageRate.HasValue && wageRate.Value >= 0.00m)

                {

                    parsedRate = wageRate.Value;

                    hasRate = true;

                }

                else

                {

                    var rawStawka = row.GetType().GetProperty("Stawka")?.GetValue(row)?.ToString();

                    if (!string.IsNullOrWhiteSpace(rawStawka) &&

                        decimal.TryParse(rawStawka.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal r) &&

                        r >= 0.00m)

                    {

                        parsedRate = r;

                        hasRate = true;

                    }

                }



                if (hasRate && parsedRate > 0)

                {

                    try

                    {

                        u.Kwota = parsedRate;

                    }

                    catch { }

                }



                // In Gratyfikant nexo, wage component validity must start on the 1st of the month

                var componentStart = new DateTime(dtStart.Year, dtStart.Month, 1);

                foreach (var sp in u.SkladnikiPlacowe)

                {

                    foreach (var val in sp.Wartosci)

                    {

                        val.Wartosc = hasRate ? parsedRate : 0.00m;

                        if (val.OkresObowiazywania == null)

                            val.OkresObowiazywania = new OkresDowolny { DataPoczatkowa = componentStart };

                        else

                            val.OkresObowiazywania.DataPoczatkowa = componentStart;

                    }

                }



                // 6. Occupation Code (Kod zawodu)

                if (!string.IsNullOrWhiteSpace(row.KodZawodu))

                {

                    try

                    {

                        var matchingZawod = context != null

                            ? context.Zawody.FirstOrDefault(z => z.Kod == row.KodZawodu)

                            : uchwyt.PodajObiektTypu<IZawody>()?.Dane?.Wszystkie()?.FirstOrDefault(z => z.Kod == row.KodZawodu);



                        if (matchingZawod != null) u.Zawod = matchingZawod;

                    }

                    catch { }

                }



                // 6b. Insurance Title Code (Kod tytułu ubezpieczenia ZUS)

                if (!string.IsNullOrWhiteSpace(row.KodTytuluUbezpieczenia))

                {

                    try

                    {

                        string targetKod = row.KodTytuluUbezpieczenia.Trim();

                        var matchingKod = context != null

                            ? context.KodyTytuluUbezpieczenia.FirstOrDefault(k => k.Kod == targetKod || k.Kod.StartsWith(targetKod))

                            : uchwyt.KodyTytuluUbezpieczenia()?.Dane?.Wszystkie()?.FirstOrDefault(k => k.Kod == targetKod || k.Kod.StartsWith(targetKod));



                        if (matchingKod != null)

                        {

                            u.KodTytuluUbezpieczenia = matchingKod;

                        }

                    }

                    catch { }

                }



                // 7. Flaga własna (OPODATKOWANE / NIEOPODATKOWANE)

                try

                {

                    string flagName = isStudentUnder26 ? "NIEOPODATKOWANE" : "OPODATKOWANE";

                    var matchingFlaga = context != null

                        ? context.FlagiWlasne.FirstOrDefault(f => f.Nazwa == flagName)

                        : uchwyt.PodajObiektTypu<IFlagiWlasne>()?.Dane?.Wszystkie()?.FirstOrDefault(f => f.Nazwa == flagName);



                    if (matchingFlaga != null) u.FlagaWlasna = matchingFlaga;

                }

                catch { }



                // 8. Save contract

                if (!umowaBO.Zapisz())

                {

                    var bledy = uchwyt.PodajBledy(umowaBO);

                    var errors = bledy != null && bledy.Any()

                        ? string.Join("; ", bledy.Select(e => $"[{e.Waznosc}] {e.Tresc}"))

                        : string.Join("; ", umowaBO.PobierzKomunikatyBledow().Select(e => e.Tresc));



                    throw new InvalidOperationException($"Błąd zapisu umowy dla {row.Imie} {row.Nazwisko}: {errors}");

                }



                return u;

            }

        }

    }

}

