param(
    [switch]$KillRunningNexo
)

$ErrorActionPreference = 'Stop'
$workspaceRoot = "c:\ NEXO SDK SFERA"

if ($KillRunningNexo) {
    Write-Host "Closing running Gratyfikant / Nexo processes to release file locks..."
    Get-Process | Where-Object { $_.ProcessName -like "*Gratyfikant*" -or $_.ProcessName -like "*Nexo*" } | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
}

$sub = Get-ChildItem -Path $workspaceRoot -Directory -Filter "HrAppka Import Prac*" | Select-Object -First 1
$projDir = $sub.FullName
$binDebugDir = Join-Path $projDir "bin\Debug"

$sourceDll = (Get-ChildItem -Path $binDebugDir -Filter "*HrAppka*.dll" | Select-Object -First 1).FullName
$sourcePdb = (Get-ChildItem -Path $binDebugDir -Filter "*HrAppka*.pdb" | Select-Object -First 1).FullName
$sourcePlugin = (Get-ChildItem -Path $binDebugDir -Filter "*HrAppka*.plugin" | Select-Object -First 1).FullName
$sourceMini = Join-Path $binDebugDir "MiniExcel.dll"

if (-not (Test-Path $sourceDll)) {
    throw "Source DLL not found in $binDebugDir"
}

$deployRoot = "C:\Users\zasc2\AppData\Local\InsERT\Deployments\Nexo"
Write-Host "Syncing compiled plugin to all Nexo deployments in $deployRoot..."

if (-not (Test-Path $deployRoot)) {
    Write-Host "WARNING: Deployments root not found: $deployRoot"
    return
}

Get-ChildItem -Path $deployRoot -Directory | ForEach-Object {
    $binDir = Join-Path $_.FullName "Binaries"
    if (Test-Path $binDir) {
        Write-Host "Copying to deployment: $($_.Name)..."
        try {
            $destDll = Join-Path $binDir (Split-Path $sourceDll -Leaf)
            [System.IO.File]::Copy($sourceDll, $destDll, $true)
            
            if ($sourcePdb -and (Test-Path $sourcePdb)) {
                $destPdb = Join-Path $binDir (Split-Path $sourcePdb -Leaf)
                [System.IO.File]::Copy($sourcePdb, $destPdb, $true)
            }
            if ($sourcePlugin -and (Test-Path $sourcePlugin)) {
                $destPlugin = Join-Path $binDir (Split-Path $sourcePlugin -Leaf)
                [System.IO.File]::Copy($sourcePlugin, $destPlugin, $true)
            }
            if (Test-Path $sourceMini) {
                $destMini = Join-Path $binDir "MiniExcel.dll"
                [System.IO.File]::Copy($sourceMini, $destMini, $true)
            }
            $info = [System.IO.FileInfo]::new($destDll)
            Write-Host "  >>> Copied successfully! Size: $($info.Length) bytes | LastWrite: $($info.LastWriteTime)"
        } catch {
            Write-Host "  WARNING: Could not overwrite DLL in $($_.Name) - process Nexo.exe is currently locking this file in memory! ($($_.Exception.Message))"
        }
    }
}

Write-Host "All deployment directories synchronized!"
