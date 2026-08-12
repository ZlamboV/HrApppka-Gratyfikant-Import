$dbs = @("Nexo_KLINEX SP_Z_O_O", "Nexo_EUROSUPPORT_GROUP", "Nexo_EUROSUPPORT OUTSOURCING")

foreach ($db in $dbs) {
    Write-Host "Installing to $db..."
    $content = "Serwer=100.127.240.115\INSERTGT
UwierzytelnianieWindows=Nie
Uzytkownik=sa
Haslo=
BazaDanych=$db"
    Set-Content -Path "ParametryInstalacji.txt" -Value $content
    dotnet run --project "c:\ NEXO SDK SFERA\HrAppka Import Pracowników\HrAppka Import Pracowników.csproj" -- --install
}
