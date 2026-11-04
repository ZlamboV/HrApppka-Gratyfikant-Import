using System;
using System.Collections.Generic;
using System.Linq;
using InsERT.Moria.Flagi;
using InsERT.Moria.Kadry.Duze;
using InsERT.Moria.Klienci;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Sfera;
using InsERT.Moria.Waluty;

namespace HrAppka_Import_Pracowników.Services
{
    public sealed record ExistingPersonInfo
    {
        public int PodmiotId { get; init; }
        public string? Pesel { get; init; }
        public string? DocumentNumber { get; init; }
        public int? DocumentType { get; init; }
        public bool IsPracownik { get; init; }
    }

    /// <summary>
    /// Reprezentuje bufor słownikowy w pamięci dla eliminacji zapytań N+1 w Sfera SDK 62.
    /// Zgodny z Wymaganiem R7 (ORIGINAL_REQUEST.md §R7, PROJECT.md §Feature 18).
    /// </summary>
    public sealed class SferaDictionaryCache
    {
        private static readonly Lazy<SferaDictionaryCache> _lazy =
            new Lazy<SferaDictionaryCache>(() => new SferaDictionaryCache());

        public static SferaDictionaryCache Instance => _lazy.Value;

        private readonly object _syncLock = new object();

        // 1. Państwa / Countries
        public Dictionary<string, Panstwo> CountriesByCode { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, Panstwo> CountriesByName { get; private set; } = new(StringComparer.OrdinalIgnoreCase);

        // 2. Obywatelstwa / Citizenships
        public Dictionary<string, Obywatelstwo> CitizenshipsByName { get; private set; } = new(StringComparer.OrdinalIgnoreCase);

        // 3. Organy Podatkowe / Tax Offices
        public Dictionary<string, OrganPodatkowy> TaxOfficesByCode { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, OrganPodatkowy> TaxOfficesByName { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
        public List<OrganPodatkowy> AllTaxOffices { get; private set; } = new();
        public List<UrzadSkarbowy> AllMasterUrzedy { get; private set; } = new();

        // 4. Działy / Departments
        public Dictionary<string, JednostkaOrganizacyjnaGr> DepartmentsByName { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
        public JednostkaOrganizacyjnaGr? RootFabrykaDepartment { get; private set; }

        // 5. Istniejący pracownicy / Existing Persons
        public Dictionary<string, ExistingPersonInfo> ExistingByPesel { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, ExistingPersonInfo> ExistingByPassport { get; private set; } = new(StringComparer.OrdinalIgnoreCase);

        // 6. Typy umów i waluty / Contract types & Currencies
        public Dictionary<string, TypUmowyPracowniczejGr> ContractTypesByName { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, Waluta> CurrenciesBySymbol { get; private set; } = new(StringComparer.OrdinalIgnoreCase);

        // 7. Flagi własne & Zawody
        public Dictionary<string, FlagaWlasna> FlagsByName { get; private set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, Zawod> ProfessionsByCode { get; private set; } = new(StringComparer.OrdinalIgnoreCase);

        public bool IsInitialized { get; private set; }

        public void Initialize(IUchwyt uchwyt)
        {
            if (uchwyt == null) throw new ArgumentNullException(nameof(uchwyt));

            lock (_syncLock)
            {
                // 1. Pre-cache Countries
                try
                {
                    var countries = uchwyt.Panstwa().Dane.Wszystkie().ToList();
                    CountriesByCode = countries
                        .Where(c => !string.IsNullOrWhiteSpace(c.KodPanstwaUE))
                        .GroupBy(c => c.KodPanstwaUE.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                    CountriesByName = countries
                        .Where(c => !string.IsNullOrWhiteSpace(c.Nazwa))
                        .GroupBy(c => c.Nazwa.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                }
                catch (Exception ex)
                {
                    LogNexo.Informacja($"[Cache] Błąd pre-cache Państwa: {ex.Message}");
                }

                // 2. Pre-cache Citizenships
                try
                {
                    var citizenships = uchwyt.Obywatelstwa().Dane.Wszystkie().ToList();
                    CitizenshipsByName = citizenships
                        .Where(c => !string.IsNullOrWhiteSpace(c.Nazwa))
                        .GroupBy(c => c.Nazwa.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                }
                catch (Exception ex)
                {
                    LogNexo.Informacja($"[Cache] Błąd pre-cache Obywatelstwa: {ex.Message}");
                }

                // 3. Pre-cache Tax Offices
                try
                {
                    var taxOfficeEntities = uchwyt.Podmioty().Dane.WszystkieOrganyPodatkowe()
                        .Where(p => p.Firma != null && p.Firma.OrganPodatkowy != null)
                        .ToList();

                    AllTaxOffices = taxOfficeEntities.Select(p => p.Firma.OrganPodatkowy).ToList();

                    TaxOfficesByCode = AllTaxOffices
                        .Where(o => !string.IsNullOrWhiteSpace(o.Kod))
                        .GroupBy(o => o.Kod.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                    TaxOfficesByName = AllTaxOffices
                        .Where(o => o.Firma != null && !string.IsNullOrWhiteSpace(o.Firma.Nazwa))
                        .GroupBy(o => o.Firma.Nazwa.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                    // Also index short names if present
                    foreach (var p in taxOfficeEntities.Where(p => !string.IsNullOrWhiteSpace(p.NazwaSkrocona)))
                    {
                        string shortName = p.NazwaSkrocona.Trim();
                        if (!TaxOfficesByName.ContainsKey(shortName))
                            TaxOfficesByName[shortName] = p.Firma.OrganPodatkowy;
                    }

                    // Pre-cache Master Template Tax Offices
                    var masterUrzedy = uchwyt.PodajObiektTypu<IUrzedySkarboweDane>();
                    if (masterUrzedy != null)
                    {
                        AllMasterUrzedy = masterUrzedy.Wszystkie().ToList();
                    }
                }
                catch (Exception ex)
                {
                    LogNexo.Informacja($"[Cache] Błąd pre-cache Organy Podatkowe: {ex.Message}");
                }

                // 4. Pre-cache Departments
                try
                {
                    var departments = uchwyt.JednostkiOrganizacyjne().Dane.Wszystkie()
                        .Where(d => !d.IsInRecycleBin)
                        .ToList();

                    DepartmentsByName = departments
                        .Where(d => !string.IsNullOrWhiteSpace(d.Nazwa))
                        .GroupBy(d => d.Nazwa.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                    RootFabrykaDepartment = departments.FirstOrDefault(d =>
                        string.Equals(d.Nazwa.Trim(), "FABRYKA", StringComparison.OrdinalIgnoreCase) && d.JednostkaNadrzedna == null);
                }
                catch (Exception ex)
                {
                    LogNexo.Informacja($"[Cache] Błąd pre-cache Działy: {ex.Message}");
                }

                // 5. Pre-cache Existing Persons & Employees (Single Query Projection!)
                try
                {
                    var existingRaw = uchwyt.Podmioty().Dane.WszystkieOsoby()
                        .Where(p => p.Osoba != null)
                        .Select(p => new
                        {
                            PodmiotId = p.Id,
                            p.Osoba.PESEL,
                            p.Osoba.SeriaINumerDokumentuTozsamosci,
                            p.Osoba.TypDokumentuTozsamosci,
                            IsPracownik = p.Osoba.Pracownik != null
                        })
                        .ToList();

                    ExistingByPesel = existingRaw
                        .Where(p => !string.IsNullOrWhiteSpace(p.PESEL))
                        .GroupBy(p => p.PESEL.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => new ExistingPersonInfo
                        {
                            PodmiotId = g.First().PodmiotId,
                            Pesel = g.Key,
                            DocumentNumber = g.First().SeriaINumerDokumentuTozsamosci,
                            DocumentType = g.First().TypDokumentuTozsamosci,
                            IsPracownik = g.First().IsPracownik
                        }, StringComparer.OrdinalIgnoreCase);

                    ExistingByPassport = existingRaw
                        .Where(p => !string.IsNullOrWhiteSpace(p.SeriaINumerDokumentuTozsamosci))
                        .GroupBy(p => p.SeriaINumerDokumentuTozsamosci.Trim(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => new ExistingPersonInfo
                        {
                            PodmiotId = g.First().PodmiotId,
                            Pesel = g.First().PESEL,
                            DocumentNumber = g.Key,
                            DocumentType = g.First().TypDokumentuTozsamosci,
                            IsPracownik = g.First().IsPracownik
                        }, StringComparer.OrdinalIgnoreCase);
                }
                catch (Exception ex)
                {
                    LogNexo.Informacja($"[Cache] Błąd pre-cache Istniejący pracownicy: {ex.Message}");
                }

                // 6. Currencies & Contract Types
                try
                {
                    var walutyMgr = uchwyt.PodajObiektTypu<IWaluty>();
                    if (walutyMgr != null)
                    {
                        CurrenciesBySymbol = walutyMgr.Dane.Wszystkie().ToList()
                            .Where(w => !string.IsNullOrWhiteSpace(w.Symbol))
                            .GroupBy(w => w.Symbol.Trim(), StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                    }

                    var typyMgr = uchwyt.TypyUmowPracowniczych();
                    if (typyMgr != null)
                    {
                        ContractTypesByName = typyMgr.Dane.Wszystkie().ToList()
                            .Where(t => !string.IsNullOrWhiteSpace(t.Nazwa))
                            .GroupBy(t => t.Nazwa.Trim(), StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                    }
                }
                catch (Exception ex)
                {
                    LogNexo.Informacja($"[Cache] Błąd pre-cache Waluty/TypyUmów: {ex.Message}");
                }

                // 7. Flags & Professions
                try
                {
                    var flagiMgr = uchwyt.PodajObiektTypu<IFlagiWlasne>();
                    if (flagiMgr != null)
                    {
                        FlagsByName = flagiMgr.Dane.Wszystkie().ToList()
                            .Where(f => !string.IsNullOrWhiteSpace(f.Nazwa))
                            .GroupBy(f => f.Nazwa.Trim(), StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                    }

                    var zawodyMgr = uchwyt.PodajObiektTypu<IZawody>();
                    if (zawodyMgr != null)
                    {
                        ProfessionsByCode = zawodyMgr.Dane.Wszystkie().ToList()
                            .Where(z => !string.IsNullOrWhiteSpace(z.Kod))
                            .GroupBy(z => z.Kod.Trim(), StringComparer.OrdinalIgnoreCase)
                            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                    }
                }
                catch (Exception ex)
                {
                    LogNexo.Informacja($"[Cache] Błąd pre-cache Flagi/Zawody: {ex.Message}");
                }

                IsInitialized = true;
            }
        }

        public Panstwo? ResolveCountry(string codeOrName)
        {
            if (string.IsNullOrWhiteSpace(codeOrName)) return null;
            string key = codeOrName.Trim();
            if (CountriesByCode.TryGetValue(key, out var p)) return p;
            if (CountriesByName.TryGetValue(key, out p)) return p;
            return null;
        }

        public OrganPodatkowy? ResolveTaxOffice(IUchwyt uchwyt, string usName)
        {
            if (string.IsNullOrWhiteSpace(usName)) return null;
            string key = usName.Trim();

            lock (_syncLock)
            {
                // 1. Direct code or name lookup
                if (TaxOfficesByCode.TryGetValue(key, out var organDirect)) return organDirect;
                if (TaxOfficesByName.TryGetValue(key, out organDirect)) return organDirect;

                // 2. Normalized name lookup on cached local offices
                string normTarget = NormalizeUsFullName(usName);
                var match = AllTaxOffices.FirstOrDefault(o =>
                    o.Firma != null && (NormalizeUsFullName(o.Firma.Nazwa) == normTarget ||
                    NormalizeUsFullName(o.Firma.Podmiot?.NazwaSkrocona ?? "") == normTarget));

                if (match != null)
                {
                    TaxOfficesByName[key] = match;
                    return match;
                }

                // 3. Token-set match on cached local offices
                var targetTokens = GetUsTokenSet(usName);
                if (targetTokens.Count > 0)
                {
                    OrganPodatkowy? bestLocal = null;
                    int bestScore = 0;

                    foreach (var o in AllTaxOffices)
                    {
                        if (o.Firma == null) continue;
                        var candTokens = GetUsTokenSet(o.Firma.Nazwa);
                        if (!string.IsNullOrWhiteSpace(o.Firma.Podmiot?.NazwaSkrocona))
                            candTokens.UnionWith(GetUsTokenSet(o.Firma.Podmiot.NazwaSkrocona));

                        int matchCount = targetTokens.Count(t => candTokens.Contains(t));
                        if (matchCount > bestScore && matchCount == targetTokens.Count)
                        {
                            bestScore = matchCount;
                            bestLocal = o;
                        }
                    }

                    if (bestLocal != null)
                    {
                        TaxOfficesByName[key] = bestLocal;
                        return bestLocal;
                    }
                }

                // 4. Fallback: Search Master Template Tax Offices and create local office if needed
                if (AllMasterUrzedy.Any())
                {
                    var exactTemplate = AllMasterUrzedy.FirstOrDefault(t =>
                        !string.IsNullOrWhiteSpace(t.Nazwa) && NormalizeUsFullName(t.Nazwa) == normTarget);

                    if (exactTemplate == null && targetTokens.Count > 0)
                    {
                        int bestScore = 0;
                        foreach (var t in AllMasterUrzedy)
                        {
                            if (string.IsNullOrWhiteSpace(t.Nazwa)) continue;
                            var candTokens = GetUsTokenSet(t.Nazwa);
                            int matchCount = targetTokens.Count(tk => candTokens.Contains(tk));
                            if (matchCount > bestScore && matchCount == targetTokens.Count)
                            {
                                bestScore = matchCount;
                                exactTemplate = t;
                            }
                        }
                    }

                    if (exactTemplate != null && !string.IsNullOrWhiteSpace(exactTemplate.Kod))
                    {
                        try
                        {
                            var podmiotyMgr = uchwyt.Podmioty();
                            using (var organBO = podmiotyMgr.UtworzOrganPodatkowy())
                            {
                                if (organBO is IOrganPodatkowy organInterface && organInterface.WypelnijDlaKodu(exactTemplate.Kod))
                                {
                                    organBO.AutoSymbol();
                                    if (organBO.Zapisz())
                                    {
                                        var createdOrgan = organBO.Dane.Firma.OrganPodatkowy;
                                        AllTaxOffices.Add(createdOrgan);
                                        TaxOfficesByCode[exactTemplate.Kod] = createdOrgan;
                                        TaxOfficesByName[key] = createdOrgan;
                                        if (!string.IsNullOrWhiteSpace(createdOrgan.Firma?.Nazwa))
                                            TaxOfficesByName[createdOrgan.Firma.Nazwa.Trim()] = createdOrgan;
                                        return createdOrgan;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            LogNexo.Informacja($"[Cache] Błąd tworzenia organu podatkowego '{exactTemplate.Nazwa}': {ex.Message}");
                        }
                    }
                }

                return null;
            }
        }

        public JednostkaOrganizacyjnaGr? ResolveDepartment(IUchwyt uchwyt, string departmentName)
        {
            if (string.IsNullOrWhiteSpace(departmentName)) return null;
            string key = departmentName.Trim();

            lock (_syncLock)
            {
                if (DepartmentsByName.TryGetValue(key, out var existing))
                    return existing;

                var deptMgr = uchwyt.JednostkiOrganizacyjne();

                // Ensure parent FABRYKA exists
                if (RootFabrykaDepartment == null)
                {
                    // Check if exists in DB
                    RootFabrykaDepartment = deptMgr.Dane.Wszystkie()
                        .FirstOrDefault(d => d.Nazwa.ToUpper() == "FABRYKA" && d.JednostkaNadrzedna == null && !d.IsInRecycleBin);

                    if (RootFabrykaDepartment == null)
                    {
                        using (var parentBO = deptMgr.Utworz())
                        {
                            parentBO.Dane.Nazwa = "FABRYKA";
                            if (parentBO.Zapisz())
                            {
                                RootFabrykaDepartment = parentBO.Dane;
                                DepartmentsByName["FABRYKA"] = RootFabrykaDepartment;
                            }
                            else
                            {
                                var bledy = string.Join("; ", parentBO.PobierzKomunikatyBledow().Select(b => b.Tresc));
                                throw new InvalidOperationException($"Nie udało się utworzyć działu nadrzędnego 'FABRYKA': {bledy}");
                            }
                        }
                    }
                    else
                    {
                        DepartmentsByName["FABRYKA"] = RootFabrykaDepartment;
                    }
                }

                // Create subdepartment
                using (var subBO = deptMgr.Utworz())
                {
                    subBO.Dane.Nazwa = key;
                    subBO.Dane.JednostkaNadrzedna = RootFabrykaDepartment;
                    if (subBO.Zapisz())
                    {
                        var created = subBO.Dane;
                        DepartmentsByName[key] = created;
                        return created;
                    }
                    else
                    {
                        var bledy = string.Join("; ", subBO.PobierzKomunikatyBledow().Select(b => b.Tresc));
                        throw new InvalidOperationException($"Nie udało się utworzyć poddziału '{key}' pod 'FABRYKA': {bledy}");
                    }
                }
            }
        }

        public bool TryGetExistingPersonByPesel(string? pesel, out ExistingPersonInfo? info)
        {
            info = null;
            if (string.IsNullOrWhiteSpace(pesel)) return false;
            return ExistingByPesel.TryGetValue(pesel.Trim(), out info);
        }

        public bool TryGetExistingPersonByPassport(string? passport, out ExistingPersonInfo? info)
        {
            info = null;
            if (string.IsNullOrWhiteSpace(passport)) return false;
            return ExistingByPassport.TryGetValue(passport.Trim(), out info);
        }

        public void RegisterNewPerson(int podmiotId, string? pesel, string? passport, bool isPracownik)
        {
            lock (_syncLock)
            {
                var entry = new ExistingPersonInfo
                {
                    PodmiotId = podmiotId,
                    Pesel = pesel?.Trim(),
                    DocumentNumber = passport?.Trim(),
                    DocumentType = null,
                    IsPracownik = isPracownik
                };

                if (!string.IsNullOrWhiteSpace(entry.Pesel))
                    ExistingByPesel[entry.Pesel] = entry;

                if (!string.IsNullOrWhiteSpace(entry.DocumentNumber))
                    ExistingByPassport[entry.DocumentNumber] = entry;
            }
        }

        public FlagaWlasna? LookupFlag(string flagName)
        {
            if (string.IsNullOrWhiteSpace(flagName)) return null;
            return FlagsByName.TryGetValue(flagName.Trim(), out var f) ? f : null;
        }

        public Zawod? LookupProfession(string kod)
        {
            if (string.IsNullOrWhiteSpace(kod)) return null;
            return ProfessionsByCode.TryGetValue(kod.Trim(), out var z) ? z : null;
        }

        private static string NormalizeUsFullName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "";
            string s = input.ToLowerInvariant();
            s = s.Replace("pierwszy", "1").Replace("i ", "1 ").Replace("drugi", "2").Replace("ii ", "2 ");
            s = s.Replace("urząd", "").Replace("urzad", "").Replace("skarbowy", "").Replace(" w ", " ");
            s = s.Replace("–", " ").Replace("-", " ").Replace(".", " ");
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ").Trim();
            return s;
        }

        private static HashSet<string> GetUsTokenSet(string input)
        {
            var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(input)) return tokens;

            var rawTokens = System.Text.RegularExpressions.Regex.Split(input.ToLowerInvariant(), @"[\s\.,\-_]+");
            var stopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "urząd", "urzad", "skarbowy", "w", "we", "us", "i", "ii", "dla" };

            foreach (var t in rawTokens)
            {
                string norm = t.Trim();
                if (norm == "pierwszy" || norm == "i") norm = "1";
                else if (norm == "drugi" || norm == "ii") norm = "2";

                if (norm.Length > 0 && !stopWords.Contains(norm))
                {
                    tokens.Add(norm);
                }
            }
            return tokens;
        }
    }
}
