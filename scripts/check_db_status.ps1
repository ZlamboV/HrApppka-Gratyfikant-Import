$ErrorActionPreference = 'Stop'
$connStr = 'Server=100.127.240.115\INSERTGT;User Id=sa;Password=;TrustServerCertificate=True;Database=Nexo_Demo_1;'
$conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
try {
    $conn.Open()
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = "SELECT DeploymentName, DeploymentManifest FROM InsLauncher.InstalledProducts"
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        $depName = $reader["DeploymentName"]
        $m = $reader["DeploymentManifest"].ToString()
        Write-Host "DeploymentName: $depName"
        if ($m -like "*HrAppka*") {
            Write-Host "  -> Matches HrAppka:"
            $m -split "`n" | Where-Object { $_ -like "*HrAppka*" } | ForEach-Object { Write-Host "     $($_.Trim())" }
        }
    }
    $reader.Close()
} catch {
    Write-Host "Error: $($_.Exception.Message)"
} finally {
    $conn.Close()
}
