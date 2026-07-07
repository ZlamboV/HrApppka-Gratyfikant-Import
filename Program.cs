using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Loader;
using System.Collections.Generic;
using InsERT.Moria.Klienci;
using InsERT.Moria.Sfera;
using InsERT.Moria.ModelDanych;
using InsERT.Mox.BusinessObjects;
using InsERT.Mox.ObiektyBiznesowe;
using InsERT.Moria.Rozszerzanie;
using MiniExcelLibs;
using HrAppka_Import_Pracowników;

class Program
{
    static Program()
    {
        AssemblyLoadContext.Default.Resolving += (context, name) =>
        {
            string assemblyName = name.Name;
            string path = Path.Combine(@"C:\ NEXO SDK SFERA\nexoSDK_60.1.1.9292\Bin", assemblyName + ".dll");
            if (File.Exists(path))
            {
                return context.LoadFromAssemblyPath(path);
            }
            return null;
        };
    }

    static void Main(string[] args)
    {
        RealMain(args);
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    static void RealMain(string[] args)
    {
        if (args.Contains("--install") || args.Contains("install"))
        {
            Console.WriteLine("=== AUTOMATIC INSTALLATION RUNNER ===");
            try
            {
                AutoInstaller.Install();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Installation Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner: {ex.InnerException.Message}");
                    Console.WriteLine(ex.InnerException.StackTrace);
                }
            }
            return;
        }

        if (args.Contains("--uninstall") || args.Contains("uninstall"))
        {
            Console.WriteLine("=== AUTOMATIC UNINSTALLATION RUNNER ===");
            try
            {
                AutoInstaller.Uninstall();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Uninstallation Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"Inner: {ex.InnerException.Message}");
                    Console.WriteLine(ex.InnerException.StackTrace);
                }
            }
            return;
        }

        Console.WriteLine("=== DIAGNOSTIC RUN OF IMPORT LOGIC ===");
        try
        {
            SferaRunner.Execute();
        }
        catch (ArgumentException argException)
        {
            LogNexo.Blad(argException);
            Console.WriteLine($"Argument Error: {argException.Message}");
            Console.WriteLine(argException.StackTrace);
        }
        catch (Exception ex)
        {
            LogNexo.Blad(ex);
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            if (ex.InnerException != null)
            {
                Console.WriteLine($"Inner: {ex.InnerException.Message}");
                Console.WriteLine(ex.InnerException.StackTrace);
            }
        }
    }
}

class AutoInstaller
{
    public static void Install()
    {
        // 1. Read ParametryInstalacji.txt
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string paramFile = Path.Combine(baseDir, "ParametryInstalacji.txt");
        if (!File.Exists(paramFile))
        {
            // Try parent directory just in case
            paramFile = Path.Combine(Directory.GetParent(baseDir)?.FullName ?? "", "ParametryInstalacji.txt");
        }
        if (!File.Exists(paramFile))
        {
            // Try project root directory
            paramFile = @"C:\ NEXO SDK SFERA\HrAppka Import Pracowników\ParametryInstalacji.txt";
        }

        if (!File.Exists(paramFile))
        {
            throw new FileNotFoundException($"Nie odnaleziono pliku parametrów instalacji: {paramFile}");
        }

        Console.WriteLine($"Reading parameters from: {paramFile}");
        var lines = File.ReadAllLines(paramFile);
        var paramsDict = lines
            .Select(l => l.Split('=', 2))
            .Where(parts => parts.Length == 2)
            .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim(), StringComparer.OrdinalIgnoreCase);

        string serwer = GetValue(paramsDict, "Serwer");
        string uwierzytelnianieWindows = GetValue(paramsDict, "UwierzytelnianieWindows");
        string uzytkownik = GetValue(paramsDict, "Uzytkownik");
        string haslo = GetValue(paramsDict, "Haslo");
        string bazaDanych = GetValue(paramsDict, "BazaDanych");

        Console.WriteLine($"Serwer: {serwer}");
        Console.WriteLine($"BazaDanych: {bazaDanych}");
        Console.WriteLine($"Uwierzytelnianie Windows: {uwierzytelnianieWindows}");
        Console.WriteLine($"Uzytkownik: {uzytkownik}");

        // Build connection strings
        var csbLauncher = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder();
        csbLauncher.DataSource = serwer;
        csbLauncher.InitialCatalog = "InsERT_Launcher";
        csbLauncher.TrustServerCertificate = true;

        if (uwierzytelnianieWindows.Equals("Tak", StringComparison.OrdinalIgnoreCase))
        {
            csbLauncher.IntegratedSecurity = true;
        }
        else
        {
            csbLauncher.IntegratedSecurity = false;
            csbLauncher.UserID = uzytkownik;
            csbLauncher.Password = haslo;
        }

        var csbPodmiot = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder(csbLauncher.ConnectionString);
        csbPodmiot.InitialCatalog = bazaDanych;

        // 2. Find package file (.mpkg)
        string pakietFolder = Path.Combine(baseDir, "Pakiet");
        if (!Directory.Exists(pakietFolder))
        {
            pakietFolder = Path.Combine(Directory.GetParent(baseDir)?.FullName ?? "", "Pakiet");
        }
        if (!Directory.Exists(pakietFolder))
        {
            pakietFolder = @"C:\ NEXO SDK SFERA\HrAppka Import Pracowników\bin\Debug\Pakiet";
        }

        if (!Directory.Exists(pakietFolder))
        {
            throw new DirectoryNotFoundException($"Nie odnaleziono katalogu pakietu: {pakietFolder}");
        }

        var mpkgFiles = Directory.GetFiles(pakietFolder, "*.mpkg");
        if (mpkgFiles.Length == 0)
        {
            throw new FileNotFoundException($"Brak pliku .mpkg w katalogu: {pakietFolder}");
        }

        // Sort to get the latest (by name or creation time)
        string mpkgPath = mpkgFiles.OrderByDescending(f => f).First();
        Console.WriteLine($"Znaleziono pakiet do instalacji: {mpkgPath}");

        // 3. Load package and upload to launcher database
        Console.WriteLine("Ładowanie pakietu...");
        var pakiet = new InsERT.Mox.Launcher.Package(mpkgPath);
        var ident = pakiet.Manifest.Identity;
        Console.WriteLine($"Pakiet: Name={ident.Name}, Version={ident.Version}");

        Console.WriteLine("Łączenie z bazą InsERT_Launcher...");
        var launchdb = new InsERT.Mox.Launcher.SqlDatabase(csbLauncher.ConnectionString);

        Console.WriteLine("Wysyłanie pakietu do bazy InsERT_Launcher...");
        bool wyslany = launchdb.UploadPackage(pakiet, true);
        Console.WriteLine($"Wynik wysłania pakietu: {(wyslany ? "Sukces" : "Pakiet już istnieje lub błąd")}");

        // 4. Update deployment manifest in the target database
        Console.WriteLine($"Łączenie z bazą podmiotu: {bazaDanych}...");
        var sqldb = new InsERT.Mox.Launcher.SqlDatabase(csbPodmiot.ConnectionString);

        Console.WriteLine("Pobieranie aktualnego manifestu wdrożenia dla Nexo...");
        var dm = sqldb.GetDeploymentManifest("Nexo");

        Console.WriteLine("Sprawdzanie czy pakiet jest już podłączony...");
        if (!dm.AdditionalPackages.Contains(ident))
        {
            //aktualizujemy manifest - usuwamy starsze wersje o tej samej nazwie przed dodaniem nowej
            var nowePakiety = dm.AdditionalPackages
                .Where(p => !p.Name.Equals(ident.Name, StringComparison.OrdinalIgnoreCase))
                .Concat(new[] { ident });

            dm = new InsERT.Mox.Launcher.DeploymentManifest(dm.Product, dm.DeploymentName, nowePakiety);
            sqldb.WriteDeploymentManifest(dm, true);
            Console.WriteLine("Manifest został pomyślnie zaktualizowany. Instalacja zakończona sukcesem!");
        }
        else
        {
            Console.WriteLine("Pakiet o tej tożsamości jest już podłączony do podmiotu. Pomijam aktualizację manifestu.");
        }
    }

    public static void Uninstall()
    {
        // 1. Read ParametryInstalacji.txt
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string paramFile = Path.Combine(baseDir, "ParametryInstalacji.txt");
        if (!File.Exists(paramFile))
        {
            paramFile = Path.Combine(Directory.GetParent(baseDir)?.FullName ?? "", "ParametryInstalacji.txt");
        }
        if (!File.Exists(paramFile))
        {
            paramFile = @"C:\ NEXO SDK SFERA\HrAppka Import Pracowników\ParametryInstalacji.txt";
        }

        if (!File.Exists(paramFile))
        {
            throw new FileNotFoundException($"Nie odnaleziono pliku parametrów instalacji: {paramFile}");
        }

        Console.WriteLine($"Reading parameters from: {paramFile}");
        var lines = File.ReadAllLines(paramFile);
        var paramsDict = lines
            .Select(l => l.Split('=', 2))
            .Where(parts => parts.Length == 2)
            .ToDictionary(parts => parts[0].Trim(), parts => parts[1].Trim(), StringComparer.OrdinalIgnoreCase);

        string serwer = GetValue(paramsDict, "Serwer");
        string uwierzytelnianieWindows = GetValue(paramsDict, "UwierzytelnianieWindows");
        string uzytkownik = GetValue(paramsDict, "Uzytkownik");
        string haslo = GetValue(paramsDict, "Haslo");
        string bazaDanych = GetValue(paramsDict, "BazaDanych");

        Console.WriteLine($"Serwer: {serwer}");
        Console.WriteLine($"BazaDanych: {bazaDanych}");
        Console.WriteLine($"Uwierzytelnianie Windows: {uwierzytelnianieWindows}");
        Console.WriteLine($"Uzytkownik: {uzytkownik}");

        // Build connection strings
        var csbLauncher = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder();
        csbLauncher.DataSource = serwer;
        csbLauncher.InitialCatalog = "InsERT_Launcher";
        csbLauncher.TrustServerCertificate = true;

        if (uwierzytelnianieWindows.Equals("Tak", StringComparison.OrdinalIgnoreCase))
        {
            csbLauncher.IntegratedSecurity = true;
        }
        else
        {
            csbLauncher.IntegratedSecurity = false;
            csbLauncher.UserID = uzytkownik;
            csbLauncher.Password = haslo;
        }

        var csbPodmiot = new Microsoft.Data.SqlClient.SqlConnectionStringBuilder(csbLauncher.ConnectionString);
        csbPodmiot.InitialCatalog = bazaDanych;

        // 2. Find package file (.mpkg) to get identity
        string pakietFolder = Path.Combine(baseDir, "Pakiet");
        if (!Directory.Exists(pakietFolder))
        {
            pakietFolder = Path.Combine(Directory.GetParent(baseDir)?.FullName ?? "", "Pakiet");
        }
        if (!Directory.Exists(pakietFolder))
        {
            pakietFolder = @"C:\ NEXO SDK SFERA\HrAppka Import Pracowników\bin\Debug\Pakiet";
        }

        if (!Directory.Exists(pakietFolder))
        {
            throw new DirectoryNotFoundException($"Nie odnaleziono katalogu pakietu: {pakietFolder}");
        }

        var mpkgFiles = Directory.GetFiles(pakietFolder, "*.mpkg");
        if (mpkgFiles.Length == 0)
        {
            throw new FileNotFoundException($"Brak pliku .mpkg w katalogu: {pakietFolder}");
        }

        string mpkgPath = mpkgFiles.OrderByDescending(f => f).First();
        Console.WriteLine($"Znaleziono pakiet: {mpkgPath}");

        var pakiet = new InsERT.Mox.Launcher.Package(mpkgPath);
        var ident = pakiet.Manifest.Identity;
        Console.WriteLine($"Pakiet do odinstalowania: Name={ident.Name}, Version={ident.Version}");

        // 3. Connect to target database and remove package from manifest
        Console.WriteLine($"Łączenie z bazą podmiotu: {bazaDanych}...");
        var sqldb = new InsERT.Mox.Launcher.SqlDatabase(csbPodmiot.ConnectionString);

        Console.WriteLine("Pobieranie aktualnego manifestu wdrożenia dla Nexo...");
        var dm = sqldb.GetDeploymentManifest("Nexo");

        Console.WriteLine("Sprawdzanie czy pakiet jest podłączony...");
        if (dm.AdditionalPackages.Contains(ident))
        {
            //aktualizujemy manifest
            dm = new InsERT.Mox.Launcher.DeploymentManifest(dm.Product, dm.DeploymentName, dm.AdditionalPackages.Except(new[] { ident }));
            sqldb.WriteDeploymentManifest(dm, true);
            Console.WriteLine("Manifest został pomyślnie zaktualizowany (pakiet usunięty z bazy).");
        }
        else
        {
            Console.WriteLine("Pakiet o tej tożsamości nie jest podłączony do podmiotu.");
        }

        // 4. Connect to launcher database and delete package from server
        Console.WriteLine("Łączenie z bazą InsERT_Launcher...");
        var launchdb = new InsERT.Mox.Launcher.SqlDatabase(csbLauncher.ConnectionString);

        Console.WriteLine("Usuwanie pakietu z serwera launcher...");
        launchdb.DeletePackage(ident);
        Console.WriteLine("Pakiet został pomyślnie usunięty z serwera launcher!");
    }

    private static string GetValue(Dictionary<string, string> dict, string key)
    {
        return dict.TryGetValue(key, out var val) ? val : "";
    }
}

class SferaRunner
{
    public static void Execute()
    {
        LogNexo.Informacja("Rozpoczęto diagnostyczny test urzędów skarbowych");

        using (InsERT.Moria.Sfera.Uchwyt uchwyt = UruchomSfere())
        {
            var testCases = new[] { 
                "Drugi Urząd Skarbowy w Opolu", 
                "Urząd Skarbowy Kraków-Krowodrza", 
                "Urząd Skarbowy Warszawa-Śródmieście" 
            };

            foreach (var testCase in testCases)
            {
                Console.WriteLine($"\n--- TESTING TAX OFFICE: {testCase} ---");
                try
                {
                    var organ = HrAppka_Import_Pracowników.OperacjaImportuPracownikow.ResolveOrganPodatkowy(uchwyt, testCase);
                    if (organ != null)
                    {
                        Console.WriteLine($"SUCCESS: Resolved to OrganPodatkowy: '{organ.Firma.Nazwa}'");
                    }
                    else
                    {
                        Console.WriteLine("FAILURE: ResolveOrganPodatkowy returned null!");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"EXCEPTION during test: {ex.Message}");
                    if (ex.InnerException != null)
                        Console.WriteLine($"  Inner Exception: {ex.InnerException.Message}");
                }
            }
        }
        LogNexo.Informacja("Zakończono diagnostyczny test urzędów skarbowych");
    }

    static void WypiszBledy(Uchwyt sfera, IObiektBiznesowy bo)
    {
        Console.WriteLine("=== ENTITY VALIDATION ERRORS ===");
        var bledy = sfera.PodajBledy(bo);
        foreach (var err in bledy)
        {
            Console.WriteLine($"  - [{err.Waznosc}] {err.Tresc}");
        }
    }

    static string GetVal(IDictionary<string, object> dict, string key)
    {
        string normKey = NormalizeKey(key);
        foreach (var k in dict.Keys)
        {
            if (NormalizeKey(k) == normKey)
            {
                return dict[k]?.ToString()?.Trim() ?? "";
            }
        }
        return "";
    }

    static string NormalizeKey(string key)
    {
        return key.ToLower()
                  .Replace("ą", "a").Replace("ć", "c").Replace("ę", "e")
                  .Replace("ł", "l").Replace("ń", "n").Replace("ó", "o")
                  .Replace("ś", "s").Replace("ź", "z").Replace("ż", "z")
                  .Replace(" ", "").Replace("-", "").Replace("/", "");
    }

    static byte GetPlec(string plecStr)
    {
        if (string.IsNullOrWhiteSpace(plecStr)) return 0;
        string lower = plecStr.ToLower();
        if (lower.Contains("kob") || lower.Contains("fem")) return 1;
        if (lower.Contains("męż") || lower.Contains("mez") || lower.Contains("mal")) return 2;
        return 0;
    }

    static Panstwo ZnajdzPanstwo(Uchwyt uchwyt, string name)
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

    static InsERT.Moria.Sfera.Uchwyt UruchomSfere()
    {
        string server = "100.127.240.115\\INSERTGT"; 
        string baza = "Nexo_Test_GR"; 
        string sqlPassword = ""; 

        var danePolaczenia = InsERT.Moria.Sfera.DanePolaczenia.Jawne(server, baza, false, "sa", sqlPassword);
        var mp = new InsERT.Moria.Sfera.MenedzerPolaczen();
        var sfera = mp.Polacz(danePolaczenia, InsERT.Mox.Product.ProductId.Gratyfikant);
        
        if (!sfera.ZalogujOperatora("Szef", "robocze"))
        {
            throw new ArgumentException("Niepoprawny operator.");
        }
        
        return sfera;
    }
}