# Available database targets:
# "Nexo_Demo_1" (active target for fast deployment)
# "Nexo_KLINEX SP_Z_O_O"
# "Nexo_EUROSUPPORT OUTSOURCING"
# "Nexo_EUROSUPPORT_GROUP"
param(
    [string[]]$Databases = @("Nexo_Demo_1", "Nexo_KLINEX SP_Z_O_O", "Nexo_EUROSUPPORT OUTSOURCING", "Nexo_EUROSUPPORT_GROUP"),
    [string]$Server = "100.127.240.115\INSERTGT",
    [switch]$QuickSyncOnly
)

$ErrorActionPreference = 'Stop'
$workspaceRoot = "c:\ NEXO SDK SFERA"
$sub = Get-ChildItem -Path $workspaceRoot -Directory -Filter "HrAppka Import Prac*" | Select-Object -First 1
$projDir = $sub.FullName
$binDebugDir = Join-Path $projDir "bin\Debug"

if ($QuickSyncOnly) {
    Write-Host ">>> QuickSyncOnly switch active: Syncing binaries directly to Nexo deployments without database reinstall <<<"
    $syncScript = Join-Path $workspaceRoot "scratch\sync_dll_to_all_deployments.ps1"
    if (Test-Path $syncScript) {
        & $syncScript
    }
    return
}

$mpkgPath = (Get-ChildItem -Path (Join-Path $projDir "Pakiet") -Filter "*.mpkg" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
if (-not $mpkgPath) {
    throw "No .mpkg found in $projDir\Pakiet"
}

Write-Host "Found MPKG release package: $mpkgPath"

$targetDbs = $Databases

foreach ($db in $targetDbs) {
    Write-Host "=========================================================="
    Write-Host "=== Installing plugin package onto database '$db' on server '$Server' ==="
    Write-Host "=========================================================="

    $paramText = @"
Serwer=$Server
UwierzytelnianieWindows=Nie
Uzytkownik=sa
Haslo=
BazaDanych=$db
"@
    $paramFileRoot = Join-Path $projDir "ParametryInstalacji.txt"
    $paramFileBin = Join-Path $binDebugDir "ParametryInstalacji.txt"

    Set-Content -Path $paramFileRoot -Value $paramText -Encoding UTF8
    if (Test-Path $binDebugDir) {
        Set-Content -Path $paramFileBin -Value $paramText -Encoding UTF8
    }

    $dllCandidate = (Get-ChildItem -Path $binDebugDir -Filter "*HrAppka*.dll" | Select-Object -First 1).FullName
    if (-not $dllCandidate) {
        $dllCandidate = Join-Path $binDebugDir "HrAppka_Import_Pracowników.dll"
    }
    Write-Host "Executing installer for '$db' with DLL '$dllCandidate'..."

    try {
        $out = & dotnet "$dllCandidate" --install 2>&1
        Write-Host ($out -join "`n")
        if ($LASTEXITCODE -eq 0 -and ($out -join "`n") -notmatch "Installation Error:") {
            Write-Host ">>> SUCCESS! Installer for '$db' finished with ExitCode 0 <<<"
        } else {
            Write-Host ">>> COMPLETED (with warnings/errors) for '$db' with ExitCode $LASTEXITCODE <<<"
        }
    } catch {
        Write-Host ">>> EXCEPTION installing on '$db': $_ <<<"
    }

    Get-Process -Name "dotnet" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like "*Import*" -or $_.MainWindowTitle -like "*HrAppka*" } | Stop-Process -Force -ErrorAction SilentlyContinue
}

# Automatically sync to local deployment folders after install so Nexo gets the fresh DLL
$syncScript = Join-Path $workspaceRoot "scratch\sync_dll_to_all_deployments.ps1"
if (Test-Path $syncScript) {
    Write-Host "Triggering deployment binary sync..."
    & $syncScript -KillRunningNexo
}

Write-Host "Target databases ($($targetDbs -join ', ')) installation and sync completed!"
