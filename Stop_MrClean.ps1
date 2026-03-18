param(
    [int]$Port = 5050
)

$conn = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
if (-not $conn) {
    Write-Host "No listening process found on port $Port"
    exit 0
}

$pids = $conn | Select-Object -ExpandProperty OwningProcess -Unique
foreach ($pid in $pids) {
    try {
        Stop-Process -Id $pid -Force
        Write-Host "Stopped process $pid on port $Port"
    } catch {
        Write-Host "Could not stop process $pid: $($_.Exception.Message)"
    }
}
