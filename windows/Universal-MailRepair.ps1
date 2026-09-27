<#
.SYNOPSIS
    Universal automated network diagnosis and repair for mail server connectivity.
.DESCRIPTION
    This script was specifically created to resolve an issue with Algerian ISP providers 
    that serve broken IPv6 addresses and randomly block certain IPv4 addresses. 
    It bypasses the local router's DNS to restore reliable connectivity.

    Features interactive menus for:
      - Test-Only Mode (no changes)
      - Repair Mode (Disables IPv6, sets custom DNS from a list, tests connectivity)
      - Rollback Mode (Restores original settings)
      - View Status Mode (Shows active physical adapters and current DNS)
.NOTES
    Author: Samir Koulali (https://samirkoulali.art)
    Version: 1.0.0
    License: Open source for non-commercial use
#>

[CmdletBinding()]
param(
    [string[]]$TargetHosts = @(),
    [switch]$TestOnly,
    [switch]$Rollback
)

# ---------------------------------------------------------------------------
# Setup: elevation check, logging, paths
# ---------------------------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "ERROR: This script must be run as Administrator. Right-click PowerShell and choose 'Run as Administrator'."
    Exit 1
}

$logDir     = Join-Path $env:ProgramData "UniversalMailRepair"
$timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile    = Join-Path $logDir "repair_$timestamp.log"
$backupFile = Join-Path $logDir "dns_backup_latest.json"

if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}

Start-Transcript -Path $logFile -Append | Out-Null
Clear-Host

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "   UNIVERSAL MAIL CONNECTION DIAGNOSTIC & REPAIR TOOL     " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Step 1: Interactive Menu - Operation Mode
# ---------------------------------------------------------------------------
$opMode = "Repair"
if ($TestOnly) { $opMode = "Test" }
elseif ($Rollback) { $opMode = "Rollback" }
else {
    $validChoice = $false
    while (-not $validChoice) {
        Write-Host "Please select an operation mode:" -ForegroundColor Yellow
        Write-Host "  1. Test Connectivity Only (Makes no changes to your network)"
        Write-Host "  2. Repair & Test (Changes DNS, Disables IPv6, Tests Connectivity)"
        Write-Host "  3. Rollback (Undo previous network changes)"
        Write-Host "  4. View Network Status (Show physical adapters and current DNS)"
        $choice = Read-Host "`nEnter 1, 2, 3, or 4 (Default: 2)"
        
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -eq "2") { $opMode = "Repair"; $validChoice = $true }
        elseif ($choice -eq "1") { $opMode = "Test"; $validChoice = $true }
        elseif ($choice -eq "3") { $opMode = "Rollback"; $validChoice = $true }
        elseif ($choice -eq "4") { $opMode = "ViewStatus"; $validChoice = $true }
        else { Write-Warning "Invalid choice. Please enter a number between 1 and 4.`n" }
    }
}

# ---------------------------------------------------------------------------
# Step 2: Interactive Menu - DNS Selection (Only if Repairing)
# ---------------------------------------------------------------------------
$selectedDns = $null
if ($opMode -eq "Repair") {
    $dnsOptions = @{
        "1" = @{ Name="Cloudflare (Recommended)"; IPs=@("1.1.1.1", "1.0.0.1") }
        "2" = @{ Name="Google"; IPs=@("8.8.8.8", "8.8.4.4") }
        "3" = @{ Name="Quad9 (Malware Blocking)"; IPs=@("9.9.9.9", "149.112.112.112") }
        "4" = @{ Name="OpenDNS"; IPs=@("208.67.222.222", "208.67.220.220") }
    }

    $validDns = $false
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    while (-not $validDns) {
        Write-Host "Select a DNS Provider to apply:" -ForegroundColor Yellow
        Write-Host "  1. $($dnsOptions['1'].Name) - [1.1.1.1 / 1.0.0.1]"
        Write-Host "  2. $($dnsOptions['2'].Name) - [8.8.8.8 / 8.8.4.4]"
        Write-Host "  3. $($dnsOptions['3'].Name) - [9.9.9.9 / 149.112.112.112]"
        Write-Host "  4. $($dnsOptions['4'].Name) - [208.67.222.222 / 208.67.220.220]"
        $dnsChoice = Read-Host "`nEnter 1, 2, 3, or 4 (Default: 1)"

        if ([string]::IsNullOrWhiteSpace($dnsChoice)) { $dnsChoice = "1" }

        if ($dnsOptions.ContainsKey($dnsChoice)) {
            $selectedDns = $dnsOptions[$dnsChoice]
            $validDns = $true
        } else {
            Write-Warning "Invalid choice. Please enter a number between 1 and 4.`n"
        }
    }
}

# ---------------------------------------------------------------------------
# Step 3: Interactive Menu - Target Host (Skip for Rollback and ViewStatus)
# ---------------------------------------------------------------------------
if ($opMode -notin @("Rollback", "ViewStatus") -and (-not $TargetHosts -or $TargetHosts.Count -eq 0)) {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    $inputHost = Read-Host "Enter the mail server hostname to test (e.g., mail.domain.com)"
    if ([string]::IsNullOrWhiteSpace($inputHost)) {
        Write-Warning "No hostname provided. Exiting."
        Stop-Transcript | Out-Null
        Exit 1
    }
    $TargetHosts = $inputHost -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
}

# ---------------------------------------------------------------------------
# Helper Function for Fast Port Testing
# ---------------------------------------------------------------------------
function Test-PortReachability {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 4000)
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    $connect = $tcpClient.BeginConnect($HostName, $Port, $null, $null)
    $success = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
    if ($success) {
        try {
            $tcpClient.EndConnect($connect)
            $tcpClient.Close()
            return $true
        } catch { return $false }
    }
    $tcpClient.Close()
    return $false
}

# ---------------------------------------------------------------------------
# EXECUTION: View Status Mode
# ---------------------------------------------------------------------------
if ($opMode -eq "ViewStatus") {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   CURRENT NETWORK & DNS STATUS                           " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    # Filter out virtual adapters, VPNs, Tailscale, WSL, and Bluetooth to keep it clean
    $physicalAdapters = Get-NetAdapter | Where-Object {
        $_.InterfaceDescription -notmatch "Hyper-V|Virtual|Tailscale|WSL|Loopback|TAP|VPN|Bluetooth|Pseudo"
    }

    if (-not $physicalAdapters) {
        Write-Warning "No physical network adapters found on this system."
    } else {
        foreach ($adapter in $physicalAdapters) {
            $statusColor = if ($adapter.Status -eq "Up") { "Green" } else { "DarkGray" }
            
            Write-Host "`nAdapter: $($adapter.Name)" -ForegroundColor Yellow
            Write-Host "Description: $($adapter.InterfaceDescription)" -ForegroundColor $statusColor
            Write-Host "Status: $($adapter.Status)" -ForegroundColor $statusColor

            # Only pull DNS info if the adapter is actually connected
            if ($adapter.Status -eq "Up") {
                $dnsConfig = Get-DnsClientServerAddress -InterfaceAlias $adapter.Name -AddressFamily IPv4 -ErrorAction SilentlyContinue
                $ipv6Binding = Get-NetAdapterBinding -Name $adapter.Name -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue
                
                $dnsServers = if ($dnsConfig.ServerAddresses) { $dnsConfig.ServerAddresses -join ', ' } else { "Automatic (DHCP / No Static DNS)" }
                $ipv6Status = if ($ipv6Binding.Enabled) { "Enabled" } else { "Disabled" }

                Write-Host "IPv4 DNS: $dnsServers" -ForegroundColor Cyan
                Write-Host "IPv6 Status: $ipv6Status" -ForegroundColor Cyan
            }
        }
    }
    
    Write-Host "`nStatus check complete." -ForegroundColor Cyan
    Stop-Transcript | Out-Null
    Exit 0
}

# ---------------------------------------------------------------------------
# EXECUTION: Rollback Mode
# ---------------------------------------------------------------------------
if ($opMode -eq "Rollback") {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    if (-not (Test-Path $backupFile)) {
        Write-Error "No backup file found at $backupFile. Nothing to roll back."
        Stop-Transcript | Out-Null
        Exit 1
    }

    Write-Host "[+] Rolling back DNS/IPv6 settings from backup..." -ForegroundColor Yellow
    $backup = @(Get-Content $backupFile -Raw | ConvertFrom-Json)

    foreach ($entry in $backup) {
        $alias = $entry.Adapter
        try {
            if ($entry.IPv6WasEnabled) {
                Enable-NetAdapterBinding -Name $alias -ComponentID ms_tcpip6 -ErrorAction Stop
                Write-Host "    [OK] Re-enabled IPv6 on '$alias'." -ForegroundColor Green
            }

            if ($entry.WasDhcp) {
                Set-DnsClientServerAddress -InterfaceAlias $alias -ResetServerAddresses -ErrorAction Stop
                Write-Host "    [OK] Restored DHCP-assigned DNS on '$alias'." -ForegroundColor Green
            } elseif ($entry.OriginalDns) {
                Set-DnsClientServerAddress -InterfaceAlias $alias -ServerAddresses $entry.OriginalDns -ErrorAction Stop
                Write-Host "    [OK] Restored original static DNS ($($entry.OriginalDns -join ', ')) on '$alias'." -ForegroundColor Green
            }
        } catch {
            Write-Warning "    [!] Failed to roll back adapter '$alias' : $_"
        }
    }

    Clear-DnsClientCache
    Write-Host "`nRollback complete." -ForegroundColor Cyan
    Stop-Transcript | Out-Null
    Exit 0
}

# ---------------------------------------------------------------------------
# EXECUTION: Repair Mode (Skip if Test-Only)
# ---------------------------------------------------------------------------
if ($opMode -eq "Repair") {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   APPLYING NETWORK FIXES                                 " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    $activeAdapters = Get-NetAdapter | Where-Object {
        $_.Status -eq "Up" -and
        $_.InterfaceDescription -notmatch "Hyper-V|Virtual|Tailscale|WSL|Loopback|TAP|VPN|Bluetooth|Pseudo"
    }

    if (-not $activeAdapters) {
        Write-Warning "No active physical network adapter detected."
        Stop-Transcript | Out-Null
        Exit 1
    }

    $backupEntries = @()

    foreach ($adapter in $activeAdapters) {
        $alias   = $adapter.Name
        $ifIndex = $adapter.InterfaceIndex
        Write-Host "[+] Processing Adapter: $alias ($($adapter.InterfaceDescription))" -ForegroundColor Yellow

        $adapterConfig    = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "InterfaceIndex = $ifIndex"
        $currentDnsConfig = Get-DnsClientServerAddress -InterfaceAlias $alias -AddressFamily IPv4 -ErrorAction SilentlyContinue
        $currentBinding   = Get-NetAdapterBinding -Name $alias -ComponentID ms_tcpip6 -ErrorAction SilentlyContinue

        $wasDhcp = [bool]($adapterConfig.DHCPEnabled -and -not $adapterConfig.DNSServerSearchOrder)

        $backupEntries += [PSCustomObject]@{
            Adapter        = $alias
            OriginalDns    = $currentDnsConfig.ServerAddresses
            WasDhcp        = $wasDhcp
            IPv6WasEnabled = [bool]($currentBinding.Enabled)
        }

        # Disable IPv6
        try {
            Disable-NetAdapterBinding -Name $alias -ComponentID ms_tcpip6 -Confirm:$false -ErrorAction Stop
            Write-Host "    [OK] IPv6 binding disabled successfully." -ForegroundColor Green
        } catch {
            Write-Warning "    [!] Could not disable IPv6 binding: $_"
        }

        # Apply chosen DNS
        try {
            Set-DnsClientServerAddress -InterfaceAlias $alias -ServerAddresses $selectedDns.IPs -ErrorAction Stop
            Write-Host "    [OK] $($selectedDns.Name) DNS ($($selectedDns.IPs -join ' / ')) configured." -ForegroundColor Green
        } catch {
            Write-Warning "    [!] Error assigning IPv4 DNS: $_"
        }
    }

    ConvertTo-Json @($backupEntries) -Depth 3 | Set-Content -Path $backupFile -Encoding UTF8
    Write-Host "`n[+] Pre-change settings backed up to: $backupFile" -ForegroundColor DarkGray
    Write-Host "    Run this script and select 'Rollback' to undo these changes." -ForegroundColor DarkGray

    Write-Host "`n[+] Flushing DNS cache..." -ForegroundColor Yellow
    Clear-DnsClientCache
    Write-Host "    [OK] DNS cache cleared." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# EXECUTION: Diagnostic Verifications (Runs for both Repair and Test modes)
# ---------------------------------------------------------------------------
Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   RUNNING CONNECTIVITY DIAGNOSTICS                       " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$allResults = @{}

foreach ($hostTarget in $TargetHosts) {
    Write-Host "`n----------------------------------------------------------"
    Write-Host "   Testing Host: $hostTarget" -ForegroundColor White
    Write-Host "----------------------------------------------------------"

    $results = [ordered]@{
        DnsResolved = "NOT TESTED"
        Imap993     = "NOT TESTED"
        Smtp465     = "NOT TESTED"
        Smtp587     = "NOT TESTED"
    }

    Write-Host "[*] Resolving DNS..." -ForegroundColor Yellow
    try {
        $dnsResult = Resolve-DnsName -Name $hostTarget -Type A -ErrorAction Stop
        $resolvedIPs = ($dnsResult | Where-Object { $_.IPAddress } | Select-Object -ExpandProperty IPAddress) -join ", "
        if ($resolvedIPs) {
            Write-Host "    [SUCCESS] Host resolved to IP: $resolvedIPs" -ForegroundColor Green
            $results.DnsResolved = "PASS"
        } else {
            Write-Host "    [FAILURE] No IPv4 A records found." -ForegroundColor Red
            $results.DnsResolved = "FAIL"
        }
    } catch {
        Write-Host "    [FAILURE] Unable to resolve $hostTarget" -ForegroundColor Red
        $results.DnsResolved = "FAIL"
    }

    Write-Host "[*] Testing IMAP (Port 993)..." -ForegroundColor Yellow
    if (Test-PortReachability -HostName $hostTarget -Port 993) {
        Write-Host "    [SUCCESS] Connected to IMAP Port 993." -ForegroundColor Green
        $results.Imap993 = "PASS"
    } else {
        Write-Host "    [FAILURE] Port 993 unreachable." -ForegroundColor Red
        $results.Imap993 = "FAIL"
    }

    Write-Host "[*] Testing SMTP (Port 465)..." -ForegroundColor Yellow
    if (Test-PortReachability -HostName $hostTarget -Port 465) {
        Write-Host "    [SUCCESS] Connected to SMTP Port 465." -ForegroundColor Green
        $results.Smtp465 = "PASS"
    } else {
        Write-Host "    [FAILURE] Port 465 unreachable." -ForegroundColor Red
        $results.Smtp465 = "FAIL"
    }

    Write-Host "[*] Testing SMTP (Port 587 - STARTTLS)..." -ForegroundColor Yellow
    if (Test-PortReachability -HostName $hostTarget -Port 587) {
        Write-Host "    [SUCCESS] Connected to SMTP Port 587." -ForegroundColor Green
        $results.Smtp587 = "PASS"
    } else {
        Write-Host "    [FAILURE] Port 587 unreachable." -ForegroundColor Red
        $results.Smtp587 = "FAIL"
    }

    $allResults[$hostTarget] = $results
}

# ---------------------------------------------------------------------------
# 3. Summary
# ---------------------------------------------------------------------------
Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   FINAL SUMMARY ($opMode Mode)                            " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

foreach ($hostTarget in $TargetHosts) {
    Write-Host "`n > Results for $hostTarget :" -ForegroundColor White
    $hostResults = $allResults[$hostTarget]
    
    foreach ($key in $hostResults.Keys) {
        $status = $hostResults[$key]
        $color = switch ($status) {
            "PASS"       { "Green" }
            "FAIL"       { "Red" }
            "NOT TESTED" { "DarkGray" }
            default      { "White" }
        }
        Write-Host ("    {0,-15} {1}" -f $key, $status) -ForegroundColor $color
    }
}

Write-Host "`nLog saved to: $logFile" -ForegroundColor DarkGray
Write-Host "Diagnostic complete." -ForegroundColor Cyan

Stop-Transcript | Out-Null
