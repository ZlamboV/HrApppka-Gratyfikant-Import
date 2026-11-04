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
            string[] searchDirs = new[]
            {
                @"c:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\Bin\",
                @"c:\ NEXO SDK SFERA\nexoSDK_61.0.4.9400\Bin\",
                @"C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo\KLINEX SP_Z_O_Oa4bfb87390e94eea9\Binaries\",
                AppDomain.CurrentDomain.BaseDirectory
            };

            foreach (var dir in searchDirs)
            {
                if (Directory.Exists(dir))
                {
                    string candidate = Path.Combine(dir, assemblyName + ".dll");
                    if (File.Exists(candidate))
                    {
                        try { return context.LoadFromAssemblyPath(candidate); } catch { }
                    }
                }
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
                Console.WriteLine("Installation finished cleanly. Exiting process.");
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
            Environment.Exit(0);
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
            string? targetFile = args.FirstOrDefault(a => !a.StartsWith("-"));
            SferaRunner.Execute(targetFile);
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

        // 4. Update deployment manifest in the target database (always replace old versions of this plugin)
        Console.WriteLine($"Łączenie z bazą podmiotu: {bazaDanych}...");
        var sqldb = new InsERT.Mox.Launcher.SqlDatabase(csbPodmiot.ConnectionString);

        Console.WriteLine("Pobieranie aktualnego manifestu wdrożenia dla Nexo...");
        var dm = sqldb.GetDeploymentManifest("Nexo");

        Console.WriteLine("Zapewnianie aktualności manifestu wdrożenia...");
        var nowePakiety = dm.AdditionalPackages
            .Where(p => !p.Name.Equals(ident.Name, StringComparison.OrdinalIgnoreCase))
            .Concat(new[] { ident });

        dm = new InsERT.Mox.Launcher.DeploymentManifest(dm.Product, dm.DeploymentName, nowePakiety);
        sqldb.WriteDeploymentManifest(dm, true);
        Console.WriteLine("Manifest został pomyślnie zaktualizowany.");

        // 5. Automatic local binaries synchronization & file-lock release
        SyncLocalDeploymentBinaries(pakietFolder);
        Console.WriteLine("Instalacja oraz synchronizacja plików zakończona sukcesem!");
    }

    private static void SyncLocalDeploymentBinaries(string sourceDllDir)
    {
        Console.WriteLine("Automatyczne zwalnianie blokad plików i synchronizacja бінаріїв...");
        try
        {
            var nexoProcs = System.Diagnostics.Process.GetProcessesByName("Nexo");
            foreach (var proc in nexoProcs)
            {
                Console.WriteLine($"Zamknięcie procesu Nexo.exe (PID {proc.Id}) w celu zwolnienia blokady plików...");
                try { proc.Kill(); proc.WaitForExit(3000); } catch { }
            }

            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string nexoDeployRoot = Path.Combine(localAppData, "InsERT", "Deployments", "Nexo");

            if (Directory.Exists(nexoDeployRoot))
            {
                foreach (var dir in Directory.GetDirectories(nexoDeployRoot))
                {
                    string binDir = Path.Combine(dir, "Binaries");
                    if (Directory.Exists(binDir))
                    {
                        string sourceDll = Path.Combine(sourceDllDir, "HrAppka_Import_Pracowników.dll");
                        string sourcePlugin = Path.Combine(sourceDllDir, "HrAppka_Import_Pracowników.plugin");
                        string sourceMini = Path.Combine(sourceDllDir, "MiniExcel.dll");

                        if (File.Exists(sourceDll)) File.Copy(sourceDll, Path.Combine(binDir, "HrAppka_Import_Pracowników.dll"), true);
                        if (File.Exists(sourcePlugin)) File.Copy(sourcePlugin, Path.Combine(binDir, "HrAppka_Import_Pracowników.plugin"), true);
                        if (File.Exists(sourceMini)) File.Copy(sourceMini, Path.Combine(binDir, "MiniExcel.dll"), true);

                        Console.WriteLine($"Pomyślnie zsynchronizowano pliki w katalogu: {Path.GetFileName(dir)}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Ostrzeżenie] Błąd synchronizacji lokalnych plików: {ex.Message}");
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
    public static void Execute(string? excelPath = null)
    {
        if (string.IsNullOrWhiteSpace(excelPath))
        {
            excelPath = @"C:\Users\zasc2\Downloads\Telegram Desktop\Eksport - Umowy - ANDROS.xlsx";
        }

        Console.WriteLine($"=== AUTONOMOUS SFERA SDK IMPORT TEST ===");
        Console.WriteLine($"Target Excel File: {excelPath}");

        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"[ERROR] Excel file not found: {excelPath}");
            return;
        }

        using (InsERT.Moria.Sfera.Uchwyt uchwyt = UruchomSfere())
        {
            Console.WriteLine("[SDK] Connected to Sfera. Reading Excel file...");
            var rows = OperacjaImportuPracownikow20.WczytajExcel(excelPath);
            Console.WriteLine($"[SDK] Loaded {rows.Count} rows from Excel.");

            Console.WriteLine("[SDK] Running pre-analysis against Nexo database...");
            OperacjaImportuPracownikow20.PreAnalyzeRows(uchwyt, rows);

            int success = 0;
            int failed = 0;

            for (int i = 0; i < rows.Count; i++)
            {
                var row = rows[i];
                Console.WriteLine($"\n--- Processing Row [{i + 1}/{rows.Count}]: {row.Imie} {row.Nazwisko} (PESEL: {row.Pesel}, Paszport: {row.Paszport}) ---");
                Console.WriteLine($"  * Adres Ulica: '{row.Ulica}', Dom: '{row.NrDomu}', Miasto: '{row.Miejscowosc}'");
                Console.WriteLine($"  * Obywatelstwo: '{row.Obywatelstwo}'");

                try
                {
                    var result = OperacjaImportuPracownikow20.ZaimportujRow20(null, row, fallbackUchwyt: uchwyt);
                    Console.WriteLine($"  * Result Status: {row.Status}");
                    if (!string.IsNullOrWhiteSpace(row.Ostrzezenia))
                    {
                        Console.WriteLine($"  * Warnings: {row.Ostrzezenia}");
                    }

                    if (row.Status == "Import zakończony sukcesem")
                    {
                        success++;
                    }
                    else
                    {
                        failed++;
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    Console.WriteLine($"  * EXCEPTION: {ex.Message}");
                    LogNexo.Blad(ex);
                }
            }

            Console.WriteLine($"\n=== TEST RUN COMPLETE ===");
            Console.WriteLine($"Total Rows: {rows.Count} | Success: {success} | Failed: {failed}");
        }
    }

    static InsERT.Moria.Sfera.Uchwyt UruchomSfere()
    {
        string server = "100.127.240.115\\INSERTGT"; 
        string baza = "Nexo_Test_1"; 
        string sqlPassword = ""; 

        string paramFile = @"C:\ NEXO SDK SFERA\HrAppka Import Pracowników\ParametryInstalacji.txt";
        if (File.Exists(paramFile))
        {
            var lines = File.ReadAllLines(paramFile);
            foreach (var line in lines)
            {
                var parts = line.Split('=', 2);
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim();
                    string val = parts[1].Trim();
                    if (key.Equals("Serwer", StringComparison.OrdinalIgnoreCase)) server = val;
                    if (key.Equals("BazaDanych", StringComparison.OrdinalIgnoreCase)) baza = val;
                    if (key.Equals("Haslo", StringComparison.OrdinalIgnoreCase)) sqlPassword = val;
                }
            }
        }

        Console.WriteLine($"[Sfera] Connecting to Server: '{server}', Database: '{baza}'...");
        var danePolaczenia = InsERT.Moria.Sfera.DanePolaczenia.Jawne(server, baza, false, "sa", sqlPassword);
        
        foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
        {
            if (asm.FullName.Contains("Moria"))
            {
                var tUzytkownicy = asm.GetType("InsERT.Moria.Uzytkownicy.IUzytkownicy");
                if (tUzytkownicy != null)
                {
                    Console.WriteLine($"[IUzytkownicy Interface] {tUzytkownicy.FullName}");
                    foreach (var m in tUzytkownicy.GetMethods())
                    {
                        var ps = string.Join(", ", Array.ConvertAll(m.GetParameters(), p => $"{p.ParameterType.Name} {p.Name}"));
                        Console.WriteLine($"   Method: {m.ReturnType.Name} {m.Name}({ps})");
                    }
                }

                var tUzytkownik = asm.GetType("InsERT.Moria.Uzytkownicy.IUzytkownik");
                if (tUzytkownik != null)
                {
                    Console.WriteLine($"[IUzytkownik Interface] {tUzytkownik.FullName}");
                    foreach (var m in tUzytkownik.GetMethods())
                    {
                        var ps = string.Join(", ", Array.ConvertAll(m.GetParameters(), p => $"{p.ParameterType.Name} {p.Name}"));
                        Console.WriteLine($"   Method: {m.ReturnType.Name} {m.Name}({ps})");
                    }
                }
            }
        }
        var mp = new InsERT.Moria.Sfera.MenedzerPolaczen();
        var sfera = mp.Polacz(danePolaczenia, InsERT.Mox.Product.ProductId.Gratyfikant);
        
        foreach (var m in sfera.GetType().GetMethods())
        {
            var ps = string.Join(", ", Array.ConvertAll(m.GetParameters(), p => p.ParameterType.Name + " " + p.Name));
            Console.WriteLine($"[Sfera Uchwyt Method] {m.ReturnType.Name} {m.Name}({ps})");
        }
        
        try
        {
            var testPracownicy = sfera.PodajObiektTypu<InsERT.Moria.Klienci.IPodmioty>();
            Console.WriteLine($"[Sfera] SUCCESS! PodajObiektTypu<IPodmioty>() returned non-null: {testPracownicy != null}");
            return sfera;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Sfera] PodajObiektTypu before Zaloguj failed: {ex.Message}");
        }

        Guid szefGuid = Guid.Parse("745a2355-ec3a-4706-9f62-f9bdec840c0e");
        
        try
        {
            if (sfera.ZalogujOperatora("Szef", ""))
            {
                Console.WriteLine("[Sfera] SUCCESS! ZalogujOperatora('Szef', '') succeeded!");
                return sfera;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Sfera] ZalogujOperatora('Szef', '') failed: {ex.Message}");
        }

        try
        {
            if (sfera.ZalogujOperatora("Szef", null))
            {
                Console.WriteLine("[Sfera] SUCCESS! ZalogujOperatora('Szef', null) succeeded!");
                return sfera;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Sfera] ZalogujOperatora('Szef', null) failed: {ex.Message}");
        }

        try
        {
            var r1 = sfera.Zaloguj("Szef", "");
            Console.WriteLine($"[Sfera] Zaloguj('Szef', '') -> {r1}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Sfera] Zaloguj('Szef', '') failed: {ex.Message}");
        }

        try
        {
            var r2 = sfera.Zaloguj("Szef", null);
            Console.WriteLine($"[Sfera] Zaloguj('Szef', null) -> {r2}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Sfera] Zaloguj('Szef', null) failed: {ex.Message}");
        }

        try
        {
            var r3 = sfera.Zaloguj(szefGuid, "");
            Console.WriteLine($"[Sfera] Zaloguj(Guid, '') -> {r3}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"[Sfera] Zaloguj(Guid, '') failed: {ex.Message}");
        }

        return sfera;
    }
}