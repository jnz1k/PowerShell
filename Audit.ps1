Import-Module ImportExcel -ErrorAction Stop

$results = @()
$serverListFile = ".\srv.txt"

if (-not (Test-Path $serverListFile)) {
    Write-Error "File $serverListFile "
    exit
}

$servers = Get-Content -Path $serverListFile

foreach ($server in $servers) {
    Write-Host "Check $server..." -ForegroundColor Cyan

    try {
        $os = Get-WmiObject Win32_OperatingSystem -ComputerName $server -ErrorAction Stop
        $computer = Get-WmiObject Win32_ComputerSystem -ComputerName $server -ErrorAction Stop
        $isHyperV = Get-WindowsFeature -ComputerName $server -Name Hyper-V -ErrorAction SilentlyContinue | Where-Object {$_.InstallState -eq "Installed"}

        $entry = [PSCustomObject]@{
            Hostname      = $server
            Manufacturer  = $computer.Manufacturer
            Model         = $computer.Model
            IsVirtual     = if ($computer.Model -like "*Virtual*") { "Yes" } else { "No" }
            OS            = $os.Caption
            HyperVRole    = if ($isHyperV) { "Yes" } else { "No" }
            VM_Count      = ""
            VM_Names      = ""
            Status        = "OK"
        }

        if ($isHyperV) {
            $vms = Invoke-Command -ComputerName $server -ScriptBlock { Get-VM } -ErrorAction SilentlyContinue
            $entry.VM_Count = $vms.Count
            $entry.VM_Names = ($vms.Name -join ", ")
        }

        $results += $entry
    }
    catch {
        $results += [PSCustomObject]@{
            Hostname      = $server
            Manufacturer  = ""
            Model         = ""
            IsVirtual     = ""
            OS            = ""
            HyperVRole    = ""
            VM_Count      = ""
            VM_Names      = ""
            Status        = "Ошибка: $_"
        }
    }
}

$excelPath = ".\result.xlsx"

$results | Export-Excel -Path $excelPath -AutoSize -BoldTopRow -FreezeTopRow -Title "Audit ann Hyper-V" `
    -ConditionalText @( 
        @{Column="HyperVRole"; Values="Yes"; BackgroundColor="LightGreen"},
        @{Column="IsVirtual"; Values="Yes"; BackgroundColor="LightSkyBlue"},
        @{Column="Status";     Values={$_ -like "Ошибка*"}; BackgroundColor="LightCoral"}
    )

Write-Host "EZX: $excelPath" -ForegroundColor Green


Get-WindowsFeature -ComputerName UN -Name Hyper-V -ErrorAction SilentlyContinue | Where-Object {$_.InstallState -eq "Installed"