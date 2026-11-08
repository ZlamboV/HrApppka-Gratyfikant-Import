using System;
using System.Collections.Concurrent;
using System.IO;
using System.Reflection;
using Microsoft.Data.Sqlite;

namespace HrAppka_Import_Pracowników.Services
{
    public sealed class AddressEnrichmentResult
    {
        public string KodPocztowy { get; init; } = string.Empty;
        public string Miejscowosc { get; init; } = string.Empty;
        public string Ulica { get; init; } = string.Empty;
        public string Gmina { get; init; } = string.Empty;
        public string Powiat { get; init; } = string.Empty;
        public string Wojewodztwo { get; init; } = string.Empty;
        public string Poczta { get; init; } = string.Empty;
    }

    /// <summary>
    /// Bezpieczny serwis wzbogacania danych adresowych (Miejscowość, Gmina, Powiat, Województwo, Poczta)
    /// w oparciu o oficjalną bazę InsERT nexo (DaneAdresowe\DaneAdresowe.db).
    /// Posiada pełny fail-safe fallback: jeśli bazy nie ma na danym komputerze, nie rzuca wyjątków.
    /// </summary>
    public sealed class AddressEnrichmentService
    {
        private static readonly Lazy<AddressEnrichmentService> _instance = new Lazy<AddressEnrichmentService>(() => new AddressEnrichmentService());
        public static AddressEnrichmentService Instance => _instance.Value;

        private readonly string? _dbPath;
        private readonly ConcurrentDictionary<string, AddressEnrichmentResult?> _cache = new ConcurrentDictionary<string, AddressEnrichmentResult?>(StringComparer.OrdinalIgnoreCase);

        private AddressEnrichmentService()
        {
            _dbPath = ResolveDatabasePath();
        }

        public bool IsDatabaseAvailable => !string.IsNullOrEmpty(_dbPath) && File.Exists(_dbPath);

        public string? DatabasePath => _dbPath;

        private static string? ResolveDatabasePath()
        {
            try
            {
                // 1. Ścieżka bazowa bieżącej domeny aplikacji (standardowy katalog binarny Nexo)
                string candidate1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DaneAdresowe", "DaneAdresowe.db");
                if (File.Exists(candidate1)) return candidate1;

                // 2. Katalog lokalizacji zmontowanej zbiorki wtyczki
                string? asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir))
                {
                    string candidate2 = Path.Combine(asmDir, "DaneAdresowe", "DaneAdresowe.db");
                    if (File.Exists(candidate2)) return candidate2;

                    // Katalog nadrzędny (jeśli wtyczka w podkatalogu)
                    string candidate3 = Path.Combine(asmDir, "..", "DaneAdresowe", "DaneAdresowe.db");
                    if (File.Exists(candidate3)) return Path.GetFullPath(candidate3);
                }

                // 3. Sprawdzenie katalogów wdrożeń Nexo w %LOCALAPPDATA%
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string nexoDeployments = Path.Combine(localAppData, "InsERT", "Deployments", "Nexo");
                if (Directory.Exists(nexoDeployments))
                {
                    foreach (var depDir in Directory.GetDirectories(nexoDeployments))
                    {
                        string cand = Path.Combine(depDir, "Binaries", "DaneAdresowe", "DaneAdresowe.db");
                        if (File.Exists(cand)) return cand;
                    }
                }
            }
            catch
            {
                // Ignorujemy błędy sprawdzania ścieżek
            }

            return null;
        }

        /// <summary>
        /// Wyszukuje wpis w DaneAdresowe.db dla podanego kodu pocztowego (format XX-XXX).
        /// Zwraca obiekt wzbogacenia lub null, jeśli nie znaleziono lub brak bazy.
        /// </summary>
        public AddressEnrichmentResult? LookupByPostalCode(string? postalCode)
        {
            if (string.IsNullOrWhiteSpace(postalCode)) return null;

            string normalizedCode = postalCode.Trim();
            if (normalizedCode.Length == 5 && !normalizedCode.Contains('-'))
            {
                normalizedCode = $"{normalizedCode.Substring(0, 2)}-{normalizedCode.Substring(2, 3)}";
            }

            if (_cache.TryGetValue(normalizedCode, out var cached))
            {
                return cached;
            }

            if (!IsDatabaseAvailable)
            {
                _cache[normalizedCode] = null;
                return null;
            }

            try
            {
                string cs = $"Data Source={_dbPath};Mode=ReadOnly";
                using var conn = new SqliteConnection(cs);
                conn.Open();

                using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT KodPocztowy, Miejscowosc, Ulica, Gmina, Powiat, Wojewodztwo FROM DaneAdresowe WHERE KodPocztowy = @code LIMIT 1";
                cmd.Parameters.AddWithValue("@code", normalizedCode);

                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    var result = new AddressEnrichmentResult
                    {
                        KodPocztowy = reader.IsDBNull(0) ? "" : reader.GetString(0),
                        Miejscowosc = reader.IsDBNull(1) ? "" : reader.GetString(1),
                        Ulica = reader.IsDBNull(2) ? "" : reader.GetString(2),
                        Gmina = reader.IsDBNull(3) ? "" : reader.GetString(3),
                        Powiat = reader.IsDBNull(4) ? "" : reader.GetString(4),
                        Wojewodztwo = reader.IsDBNull(5) ? "" : reader.GetString(5),
                        Poczta = reader.IsDBNull(1) ? "" : reader.GetString(1)
                    };
                    _cache[normalizedCode] = result;
                    return result;
                }
            }
            catch
            {
                // W razie problemu z bazą, działamy bez błędu
            }

            _cache[normalizedCode] = null;
            return null;
        }

        /// <summary>
        /// Uzupełnia brakujące pola adresu (Miejscowość, Gmina, Powiat, Województwo, Poczta)
        /// w obiekcie PracownikImportRow, jeśli nie zostały podane wejściowo.
        /// </summary>
        public void EnrichRowAddress(PracownikImportRow row)
        {
            if (row == null || string.IsNullOrWhiteSpace(row.KodPocztowy)) return;

            var info = LookupByPostalCode(row.KodPocztowy);
            if (info == null) return;

            if (string.IsNullOrWhiteSpace(row.Miejscowosc) && !string.IsNullOrWhiteSpace(info.Miejscowosc))
            {
                row.Miejscowosc = info.Miejscowosc;
            }

            if (string.IsNullOrWhiteSpace(row.Gmina) && !string.IsNullOrWhiteSpace(info.Gmina))
            {
                row.Gmina = info.Gmina;
            }

            if (string.IsNullOrWhiteSpace(row.Powiat) && !string.IsNullOrWhiteSpace(info.Powiat))
            {
                row.Powiat = info.Powiat;
            }

            if (string.IsNullOrWhiteSpace(row.Wojewodztwo) && !string.IsNullOrWhiteSpace(info.Wojewodztwo))
            {
                row.Wojewodztwo = info.Wojewodztwo;
            }
        }
    }
}
