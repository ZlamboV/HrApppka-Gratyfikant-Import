$ErrorActionPreference = 'Stop'
$workspaceRoot = "c:\ NEXO SDK SFERA"
$sub = Get-ChildItem -Path $workspaceRoot -Directory -Filter "HrAppka Import Prac*" | Select-Object -First 1
$projDir = $sub.FullName
$binDir = Join-Path $projDir "bin\Debug"

Write-Host "Copying build outputs to root for packaging from $binDir..."
Get-ChildItem -Path $binDir -Filter "*Import_Pracownik*.dll" | Copy-Item -Destination $projDir -Force
Get-ChildItem -Path $binDir -Filter "*Import_Pracownik*.plugin" | Copy-Item -Destination $projDir -Force
Get-ChildItem -Path $binDir -Filter "MiniExcel.dll" | Copy-Item -Destination $projDir -Force

$builderExe = "C:\ NEXO SDK SFERA\nexoSDK_62.0.0.9537\Narzedzia\PackageAndInstallerBuilder.exe"
$configXml = Join-Path $projDir "Konfiguracja.xml"

Write-Host "Running PackageAndInstallerBuilder.exe..."
$env:APPDATA = "c:\ NEXO SDK SFERA"
$env:GIT_CONFIG_GLOBAL = "c:\ NEXO SDK SFERA\.gitconfig"
$env:DOTNET_CLI_TELEMETRY_OPTOUT = "1"
$env:DOTNET_NOLOGO = "1"
Push-Location $projDir
try {
    & $builderExe "/PlikKonfiguracji:$configXml"
} finally {
    Pop-Location
}

# Baseline artifact definition: 1.0.8.14
$mpkgItem = Get-ChildItem -Path (Join-Path $projDir "Pakiet") -Filter "*.mpkg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$exeItem = Get-ChildItem -Path (Join-Path $projDir "Instalator") -Filter "*.exe" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if (-not $exeItem) {
    $binExe = Get-ChildItem -Path $binDir -Filter "*Import_Pracownik*.exe" | Select-Object -First 1
    if ($binExe) {
        $targetExe = Join-Path (Join-Path $projDir "Instalator") $binExe.Name
        Copy-Item -Path $binExe.FullName -Destination $targetExe -Force
        $exeItem = Get-Item $targetExe
    }
}

if ($mpkgItem) {
    Write-Host ">>> SUCCESS! Package created: $($mpkgItem.FullName) (Size: $($mpkgItem.Length) bytes) <<<"
} else {
    Write-Host "ERROR: Package file not found in Pakiet"
}

if ($exeItem) {
    Write-Host ">>> SUCCESS! Installer created: $($exeItem.FullName) (Size: $($exeItem.Length) bytes) <<<"
} else {
    Write-Host "NOTE: Installer EXE not found in Instalator"
}