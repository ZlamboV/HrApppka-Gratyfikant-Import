using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using InsERT.Moria.Klienci;
using InsERT.Moria.ModelDanych;
using InsERT.Moria.Rozszerzanie;
using InsERT.Moria.Sfera;

namespace HrAppka_Import_Pracowników.Services
{
    /// <summary>
    /// Reprezentuje pojedynczy wpis słownikowy państwa i obywatelstwa zgodny ze specyfikacją ZUS.
    /// </summary>
    public sealed record ZusCountryEntry
    {
        public int Id { get; init; }
        public string PolishName { get; init; } = string.Empty;
        public string IsoCode { get; init; } = string.Empty;
        public string ZusAdjective { get; init; } = string.Empty;
    }

    /// <summary>
    /// Menadżer słownika obywatelstw ZUS wczytywanych w formacie UTF-8 z pliku Obywatelstwa_ZUS.csv.
    /// Rozwiązuje problem błędnego przypisywania polskiego obywatelstwa cudzoziemcom,
    /// obsługując poprawne znaki diakrytyczne oraz brak cichego domyślnego powrotu do Polski (PL).
    /// </summary>
    public sealed class ZusCitizenshipManager
    {
        private static readonly Lazy<ZusCitizenshipManager> _lazyInstance =
            new Lazy<ZusCitizenshipManager>(() => new ZusCitizenshipManager());

        public static ZusCitizenshipManager Instance => _lazyInstance.Value;

        private readonly List<ZusCountryEntry> _entries = new List<ZusCountryEntry>();
        private readonly Dictionary<string, ZusCountryEntry> _byIso = new Dictionary<string, ZusCountryEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ZusCountryEntry> _byAdjective = new Dictionary<string, ZusCountryEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ZusCountryEntry> _byPolishName = new Dictionary<string, ZusCountryEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ZusCountryEntry> _byNormalizedAscii = new Dictionary<string, ZusCountryEntry>(StringComparer.OrdinalIgnoreCase);

        private readonly object _lock = new object();
        public bool IsInitialized { get; private set; }
        public int EntryCount => _entries.Count;

        /// <summary>
        /// Opcjonalne delegaty logowania dla testów jednostkowych i środowisk zewnętrznych.
        /// </summary>
        public static Action<string>? ExternalLogger { get; set; }
        public static Action<Exception>? ExternalErrorLogger { get; set; }

        // 27 państw członkowskich Unii Europejskiej
        private static readonly HashSet<string> EuIsoCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "AT", "BE", "BG", "HR", "CY", "CZ", "DK", "EE", "FI", "FR",
            "DE", "GR", "HU", "IE", "IT", "LV", "LT", "LU", "MT", "NL",
            "PL", "PT", "RO", "SK", "SI", "ES", "SE"
        };

        public ZusCitizenshipManager()
        {
            Initialize();
        }

        public void Initialize(string? explicitCsvPath = null)
        {
            lock (_lock)
            {
                if (IsInitialized && string.IsNullOrEmpty(explicitCsvPath))
                    return;

                _entries.Clear();
                _byIso.Clear();
                _byAdjective.Clear();
                _byPolishName.Clear();
                _byNormalizedAscii.Clear();

                using var stream = OpenCsvStream(explicitCsvPath);
                if (stream == null)
                {
                    LogInfo("[ZusCitizenshipManager] Ostrzeżenie: Nie odnaleziono pliku Obywatelstwa_ZUS.csv.");
                    return;
                }

                using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
                string? header = reader.ReadLine(); // Skip: Lp,Kraj_PL,Kod_ISO,Obywatelstwo_ZUS

                string? line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var parts = ParseCsvLine(line);
                    if (parts.Count >= 4)
                    {
                        if (int.TryParse(parts[0], out int id))
                        {
                            var entry = new ZusCountryEntry
                            {
                                Id = id,
                                PolishName = parts[1].Trim(),
                                IsoCode = parts[2].Trim().ToUpperInvariant(),
                                ZusAdjective = parts[3].Trim().ToLowerInvariant()
                            };

                            _entries.Add(entry);

                            if (!_byIso.ContainsKey(entry.IsoCode))
                                _byIso[entry.IsoCode] = entry;

                            if (!_byAdjective.ContainsKey(entry.ZusAdjective))
                                _byAdjective[entry.ZusAdjective] = entry;

                            if (!_byPolishName.ContainsKey(entry.PolishName))
                                _byPolishName[entry.PolishName] = entry;

                            string normAdj = RemoveDiacritics(entry.ZusAdjective);
                            if (!_byNormalizedAscii.ContainsKey(normAdj))
                                _byNormalizedAscii[normAdj] = entry;

                            string normName = RemoveDiacritics(entry.PolishName);
                            if (!_byNormalizedAscii.ContainsKey(normName))
                                _byNormalizedAscii[normName] = entry;
                        }
                    }
                }

                IsInitialized = _entries.Count > 0;
                LogInfo($"[ZusCitizenshipManager] Pomyślnie załadowano {_entries.Count} państw i obywatelstw ZUS w standardzie UTF-8.");
            }
        }

        private Stream? OpenCsvStream(string? explicitCsvPath)
        {
            if (!string.IsNullOrEmpty(explicitCsvPath) && File.Exists(explicitCsvPath))
                return File.OpenRead(explicitCsvPath);

            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string localFile = Path.Combine(baseDir, "Obywatelstwa_ZUS.csv");
            if (File.Exists(localFile))
                return File.OpenRead(localFile);

            string? asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (!string.IsNullOrEmpty(asmDir))
            {
                string asmFile = Path.Combine(asmDir, "Obywatelstwa_ZUS.csv");
                if (File.Exists(asmFile))
                    return File.OpenRead(asmFile);
            }

            const string devPath = @"c:\ NEXO SDK SFERA\HrAppka Import Pracowników\Obywatelstwa_ZUS.csv";
            if (File.Exists(devPath))
                return File.OpenRead(devPath);

            // Pobranie z zasobu wkomponowanego w bibliotekę (EmbeddedResource)
            var asm = Assembly.GetExecutingAssembly();
            var resourceName = asm.GetManifestResourceNames()
                .FirstOrDefault(n => n.EndsWith("Obywatelstwa_ZUS.csv", StringComparison.OrdinalIgnoreCase));
            if (resourceName != null)
                return asm.GetManifestResourceStream(resourceName);

            return null;
        }

        public ZusCountryEntry? MatchEntry(string? rawInput)
        {
            if (string.IsNullOrWhiteSpace(rawInput)) return null;
            if (!IsInitialized) Initialize();

            string norm = rawInput.Trim();

            // 1. Dokładne dopasowanie do przymiotnika ZUS ("ukraińskie", "polskie")
            if (_byAdjective.TryGetValue(norm, out var entry))
                return entry;

            // 2. Dokładne dopasowanie do polskiej nazwy państwa ("Ukraina", "Polska")
            if (_byPolishName.TryGetValue(norm, out entry))
                return entry;

            // 3. Dokładne dopasowanie do kodu ISO ("UA", "PL")
            if (_byIso.TryGetValue(norm, out entry))
                return entry;

            // 3a. Obsługa potocznych rzeczowników narodowościowych (Polak, Polka)
            if (norm.Equals("Polak", StringComparison.OrdinalIgnoreCase) ||
                norm.Equals("Polka", StringComparison.OrdinalIgnoreCase))
            {
                if (_byIso.TryGetValue("PL", out entry))
                    return entry;
            }

            // 3b. Obsługa formatów złożonych z separatorem (np. "PL - Polska", "UA - Ukraina", "Ukraina / UA")
            if (norm.Contains('-') || norm.Contains('/') || norm.Contains('(') || norm.Contains(':'))
            {
                var tokens = norm.Split(new[] { '-', '/', '(', ')', ':', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var token in tokens)
                {
                    string t = token.Trim();
                    if (_byIso.TryGetValue(t, out entry)) return entry;
                    if (_byPolishName.TryGetValue(t, out entry)) return entry;
                    if (_byAdjective.TryGetValue(t, out entry)) return entry;
                }
            }

            // 4. Dopasowanie prefiksowe / podciągów (np. "Ukraińskie (Ukraina)" lub "Ukraina - UA")
            entry = _entries.FirstOrDefault(e =>
                norm.StartsWith(e.ZusAdjective, StringComparison.OrdinalIgnoreCase) ||
                norm.StartsWith(e.PolishName, StringComparison.OrdinalIgnoreCase) ||
                (norm.Length >= 4 && e.PolishName.StartsWith(norm, StringComparison.OrdinalIgnoreCase)) ||
                (norm.Length >= 4 && e.ZusAdjective.StartsWith(norm, StringComparison.OrdinalIgnoreCase))
            );
            if (entry != null) return entry;

            // 5. Dopasowanie bez znaków diakrytycznych ("ukrainskie" -> "ukraińskie")
            string ascii = RemoveDiacritics(norm);
            if (_byNormalizedAscii.TryGetValue(ascii, out entry))
                return entry;

            // ŚCISŁA REGUŁA: Nigdy nie zwracamy domyślnie Polski dla nierozpoznanych wartości!
            return null;
        }

        /// <summary>
        /// Zgodne z kontraktem interfejsowym w PROJECT.md:
        /// ZusCitizenshipManager.Resolve(string nationalityString) -> (Panstwo panstwo, Obywatelstwo obywatelstwo, string kodIso)
        /// </summary>
        public (Panstwo? panstwo, Obywatelstwo? obywatelstwo, string? kodIso) Resolve(IUchwyt uchwyt, string? nationalityString)
        {
            var (panstwo, oby, entry) = ResolveSfera(uchwyt, nationalityString);
            return (panstwo, oby, entry?.IsoCode);
        }

        public (Panstwo? Panstwo, Obywatelstwo? Obywatelstwo, ZusCountryEntry? Entry) ResolveSfera(
            IUchwyt uchwyt, 
            string? rawInput)
        {
            var entry = MatchEntry(rawInput);
            if (entry == null)
            {
                LogInfo($"[ZusCitizenshipManager] Nie rozpoznano obywatelstwa dla wartości: '{rawInput}'. Pomijanie automatycznego przypisania.");
                return (null, null, null);
            }

            var panstwaMgr = uchwyt.PodajObiektTypu<IPanstwa>();
            var panstwo = panstwaMgr.Dane.Wszystkie().FirstOrDefault(p =>
                p.KodPanstwaUE == entry.IsoCode ||
                p.Nazwa.Equals(entry.PolishName, StringComparison.OrdinalIgnoreCase)
            );

            if (panstwo == null && !string.IsNullOrEmpty(entry.IsoCode))
            {
                try
                {
                    using (var panstwoBO = panstwaMgr.Utworz())
                    {
                        panstwoBO.Dane.Nazwa = entry.PolishName;
                        panstwoBO.Dane.KodPanstwaUE = entry.IsoCode;
                        panstwoBO.Dane.CzlonekUE = EuIsoCodes.Contains(entry.IsoCode);

                        if (panstwoBO.Zapisz())
                        {
                            panstwo = panstwoBO.Dane;
                            LogInfo($"[ZusCitizenshipManager] Utworzono brakujące państwo w Nexo: {entry.PolishName} ({entry.IsoCode})");
                        }
                        else
                        {
                            var errs = uchwyt.PodajBledy(panstwoBO);
                            LogInfo($"[ZusCitizenshipManager] Błąd zapisu państwa {entry.PolishName}: {string.Join("; ", errs.Select(e => e.Tresc))}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError(ex);
                }
            }

            if (panstwo == null)
                return (null, null, entry);

            var obyMgr = uchwyt.PodajObiektTypu<IObywatelstwa>();
            var obywatelstwo = obyMgr.Dane.Wszystkie().FirstOrDefault(o =>
                (o.Panstwo != null && o.Panstwo.Id == panstwo.Id && o.Nazwa.Equals(entry.ZusAdjective, StringComparison.OrdinalIgnoreCase)) ||
                o.Nazwa.Equals(entry.ZusAdjective, StringComparison.OrdinalIgnoreCase) ||
                o.Nazwa.Equals(entry.PolishName, StringComparison.OrdinalIgnoreCase)
            );

            if (obywatelstwo == null)
            {
                try
                {
                    using (var obyBO = obyMgr.Utworz())
                    {
                        obyBO.Dane.Nazwa = entry.ZusAdjective;
                        obyBO.Dane.Panstwo = panstwo;

                        if (obyBO.Zapisz())
                        {
                            obywatelstwo = obyBO.Dane;
                            LogInfo($"[ZusCitizenshipManager] Utworzono brakujące obywatelstwo: {entry.ZusAdjective}");
                        }
                        else
                        {
                            var errs = uchwyt.PodajBledy(obyBO);
                            LogInfo($"[ZusCitizenshipManager] Błąd zapisu obywatelstwa {entry.ZusAdjective}: {string.Join("; ", errs.Select(e => e.Tresc))}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogError(ex);
                }
            }

            return (panstwo, obywatelstwo, entry);
        }

        public Panstwo? ResolvePanstwo(IUchwyt uchwyt, string? inputName, out string platnikCitizenshipName)
        {
            platnikCitizenshipName = string.Empty;
            if (string.IsNullOrWhiteSpace(inputName)) return null;

            var (panstwo, _, entry) = ResolveSfera(uchwyt, inputName);
            if (entry != null)
            {
                platnikCitizenshipName = entry.ZusAdjective;
            }
            return panstwo;
        }

        public Obywatelstwo? ResolveObywatelstwo(IUchwyt uchwyt, string? inputName)
        {
            if (string.IsNullOrWhiteSpace(inputName)) return null;
            var (_, oby, _) = ResolveSfera(uchwyt, inputName);
            return oby;
        }

        public static bool IsEuMember(string? isoCode)
        {
            if (string.IsNullOrWhiteSpace(isoCode)) return false;
            return EuIsoCodes.Contains(isoCode.Trim());
        }

        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
            result.Add(sb.ToString());
            return result;
        }

        private static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            string normalized = text
                .Replace("ł", "l").Replace("Ł", "L");

            string formD = normalized.Normalize(NormalizationForm.FormD);
            var sb = new StringBuilder();

            foreach (char ch in formD)
            {
                var uc = CharUnicodeInfo.GetUnicodeCategory(ch);
                if (uc != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(ch);
                }
            }

            return sb.ToString().Normalize(NormalizationForm.FormC).ToLowerInvariant();
        }

        private static void LogInfo(string message)
        {
            if (ExternalLogger != null)
            {
                try { ExternalLogger(message); return; } catch { }
            }
            SferaLoggerBridge.Informacja(message);
        }

        private static void LogError(Exception ex)
        {
            if (ExternalErrorLogger != null)
            {
                try { ExternalErrorLogger(ex); return; } catch { }
            }
            SferaLoggerBridge.Blad(ex);
        }

        /// <summary>
        /// Mostek bezpiecznego logowania odcinający twardą zależność JIT od InsERT.Moria.Sfera.
        /// W środowisku Sfery wiąże natywny delegat do LogNexo. W testach jednostkowych przekierowuje do Trace.
        /// </summary>
        private static class SferaLoggerBridge
        {
            private static readonly Action<string>? _logInfoAction;
            private static readonly Action<Exception>? _logErrorAction;

            static SferaLoggerBridge()
            {
                try
                {
                    Type? logType = null;
                    foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
                    {
                        try
                        {
                            logType = asm.GetType("InsERT.Moria.Rozszerzanie.Operacje.LogNexo");
                            if (logType != null) break;
                        }
                        catch { }
                    }

                    if (logType == null)
                    {
                        try
                        {
                            var asm = Assembly.Load("InsERT.Moria.Sfera");
                            logType = asm?.GetType("InsERT.Moria.Rozszerzanie.Operacje.LogNexo");
                        }
                        catch { }
                    }

                    if (logType != null)
                    {
                        var infoMethod = logType.GetMethod("Informacja", new[] { typeof(string) });
                        if (infoMethod != null)
                        {
                            _logInfoAction = (Action<string>)Delegate.CreateDelegate(typeof(Action<string>), infoMethod);
                        }

                        var errorMethod = logType.GetMethod("Blad", new[] { typeof(Exception) });
                        if (errorMethod != null)
                        {
                            _logErrorAction = (Action<Exception>)Delegate.CreateDelegate(typeof(Action<Exception>), errorMethod);
                        }
                    }
                }
                catch
                {
                    // Izolacja środowiska testowego od środowiska Sfery
                }
            }

            public static void Informacja(string message)
            {
                try
                {
                    if (_logInfoAction != null)
                    {
                        _logInfoAction(message);
                        return;
                    }
                }
                catch { }

                System.Diagnostics.Trace.WriteLine(message);
            }

            public static void Blad(Exception ex)
            {
                try
                {
                    if (_logErrorAction != null)
                    {
                        _logErrorAction(ex);
                        return;
                    }
                }
                catch { }

                System.Diagnostics.Trace.WriteLine(ex.ToString());
            }
        }
    }
}
