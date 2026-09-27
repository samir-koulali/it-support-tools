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
      - Comprehensive DNS Inspector (A, AAAA, CNAME, MX, TXT/SPF, DMARC, NS)
      - Hosting, Web & Mail Port Diagnostics (IMAP, POP3, SMTP, Webmail, cPanel/N0C, HTTPS)
      - SSL / TLS Certificate Validity & SAN verification
      - One-Click Clipboard Export of diagnostic report
.NOTES
    Author: Samir Koulali (https://samirkoulali.art)
    Version: 1.1.0
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
    try { Stop-Transcript | Out-Null } catch {}; Exit 1
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
Write-Host "   UNIVERSAL MAIL & SERVER DIAGNOSTIC & REPAIR TOOL       " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# ---------------------------------------------------------------------------
# Helper: Fast TCP Port Reachability Check
# ---------------------------------------------------------------------------
function Test-PortReachability {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 3500)
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    try {
        $connect = $tcpClient.BeginConnect($HostName, $Port, $null, $null)
        $success = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($success) {
            $tcpClient.EndConnect($connect)
            $tcpClient.Close()
            return $true
        }
    } catch {
        # ignore error, returns false
    } finally {
        $tcpClient.Close()
    }
    return $false
}

# ---------------------------------------------------------------------------
# Helper: SSL / TLS Certificate Verification
# ---------------------------------------------------------------------------
function Test-SslCertificate {
    param([string]$HostName, [int]$Port = 443, [int]$TimeoutMs = 4000)
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    try {
        $connect = $tcpClient.BeginConnect($HostName, $Port, $null, $null)
        if (-not $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            $tcpClient.Close()
            return [PSCustomObject]@{ Valid = $false; Error = "Connection timed out" }
        }
        $tcpClient.EndConnect($connect)

        $sslStream = New-Object System.Net.Security.SslStream(
            $tcpClient.GetStream(),
            $false,
            ({ $true } -as [System.Net.Security.RemoteCertificateValidationCallback])
        )

        $sslStream.AuthenticateAsClient($HostName)
        $remoteCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($sslStream.RemoteCertificate)

        $now = Get-Date
        $daysLeft = [math]::Round(($remoteCert.NotAfter - $now).TotalDays)
        $isExpired = $now -gt $remoteCert.NotAfter

        $sslStream.Close()
        $tcpClient.Close()

        return [PSCustomObject]@{
            Valid       = (-not $isExpired)
            Subject     = $remoteCert.Subject
            Issuer      = $remoteCert.Issuer
            ExpiresOn   = $remoteCert.NotAfter.ToString("yyyy-MM-dd")
            DaysLeft    = $daysLeft
            SAN         = ($remoteCert.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" } | ForEach-Object { $_.Format($true) })
            Error       = if ($isExpired) { "Certificate Expired!" } else { $null }
        }
    } catch {
        return [PSCustomObject]@{ Valid = $false; Error = $_.Exception.Message }
    } finally {
        $tcpClient.Close()
    }
}

# ---------------------------------------------------------------------------
# Helper: DNS Full Record Inspector
# ---------------------------------------------------------------------------
function Show-FullDnsRecords {
    param([string]$Domain)

    # Normalize: strip "mail." or "webmail." to get the base domain for MX/TXT/DMARC checks
    $baseDomain = $Domain
    if ($Domain -match "^(mail|webmail|smtp|imap|pop|autodiscover|autoconfig|cpanel|whm)\.(.+\..+)$") {
        $baseDomain = $Matches[2]
    }

    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   FULL DNS INSPECTOR: $Domain (Base: $baseDomain)        " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

    # 1. A & AAAA Records for the exact target host
    Write-Host "`n[*] A / AAAA Records for $($Domain):" -ForegroundColor Yellow
    try {
        $aRecords = Resolve-DnsName -Name $Domain -Type A -ErrorAction Stop
        foreach ($r in $aRecords) {
            Write-Host "    A     -> $($r.IPAddress) (TTL: $($r.TTL)s)" -ForegroundColor Green
        }
    } catch {
        Write-Host "    A     -> [NONE / Failed to resolve]" -ForegroundColor Red
    }

    try {
        $aaaaRecords = Resolve-DnsName -Name $Domain -Type AAAA -ErrorAction Stop
        foreach ($r in $aaaaRecords) {
            Write-Host "    AAAA  -> $($r.IP6Address) (TTL: $($r.TTL)s)" -ForegroundColor Green
        }
    } catch {
        Write-Host "    AAAA  -> [NONE or Disabled]" -ForegroundColor DarkGray
    }

    # 2. CNAME Records
    try {
        $cname = Resolve-DnsName -Name $Domain -Type CNAME -ErrorAction Stop
        foreach ($c in $cname) {
            Write-Host "    CNAME -> $($c.NameHost)" -ForegroundColor Green
        }
    } catch { }

    # 3. MX Records (Base Domain)
    Write-Host "`n[*] MX (Mail Exchanger) Records for $($baseDomain):" -ForegroundColor Yellow
    try {
        $mxRecords = Resolve-DnsName -Name $baseDomain -Type MX -ErrorAction Stop | Sort-Object Preference
        foreach ($mx in $mxRecords) {
            Write-Host "    MX (Pref: $($mx.Preference)) -> $($mx.NameExchange)" -ForegroundColor Green
        }
    } catch {
        Write-Host "    MX    -> [No MX record found on $baseDomain]" -ForegroundColor Red
    }

    # 4. Nameservers (NS Records)
    Write-Host "`n[*] Nameservers (NS) for $($baseDomain):" -ForegroundColor Yellow
    try {
        $nsRecords = Resolve-DnsName -Name $baseDomain -Type NS -ErrorAction Stop
        foreach ($ns in $nsRecords) {
            Write-Host "    NS    -> $($ns.NameHost)" -ForegroundColor Green
        }
    } catch {
        Write-Host "    NS    -> [Could not resolve Nameservers]" -ForegroundColor Red
    }

    # 5. SPF (TXT Records on Base Domain)
    Write-Host "`n[*] SPF & TXT Records for $($baseDomain):" -ForegroundColor Yellow
    try {
        $txtRecords = Resolve-DnsName -Name $baseDomain -Type TXT -ErrorAction Stop
        $foundSpf = $false
        foreach ($txt in $txtRecords) {
            $txtVal = ($txt.Strings -join "")
            if ($txtVal -like "v=spf1*") {
                Write-Host "    SPF   -> $txtVal" -ForegroundColor Green
                $foundSpf = $true
            } else {
                Write-Host "    TXT   -> $txtVal" -ForegroundColor DarkGray
            }
        }
        if (-not $foundSpf) {
            Write-Host "    [WARNING] No v=spf1 record detected for $baseDomain!" -ForegroundColor Red
        }
    } catch {
        Write-Host "    TXT   -> [No TXT records found]" -ForegroundColor DarkGray
    }

    # 6. DMARC (_dmarc.<baseDomain>)
    Write-Host "`n[*] DMARC Record for _dmarc.$($baseDomain):" -ForegroundColor Yellow
    try {
        $dmarc = Resolve-DnsName -Name "_dmarc.$baseDomain" -Type TXT -ErrorAction Stop
        $dmarcVal = ($dmarc.Strings -join "")
        Write-Host "    DMARC -> $dmarcVal" -ForegroundColor Green
    } catch {
        Write-Host "    DMARC -> [No DMARC record configured (_dmarc.$baseDomain)]" -ForegroundColor Red
    }

    # 7. Autodiscover / Autoconfig CNAME
    Write-Host "`n[*] Mail Client Auto-Discovery Records:" -ForegroundColor Yellow
    foreach ($autoSub in @("autodiscover.$baseDomain", "autoconfig.$baseDomain")) {
        try {
            $autoRec = Resolve-DnsName -Name $autoSub -Type CNAME -ErrorAction Stop
            Write-Host "    $autoSub -> $($autoRec.NameHost)" -ForegroundColor Green
        } catch {
            Write-Host "    $autoSub -> [Not configured]" -ForegroundColor DarkGray
        }
    }
}

$isInteractive = (-not $TestOnly -and -not $Rollback)

while ($true) {
    if ($isInteractive) { 
        $TargetHosts = @() 
        Write-Host "`n==========================================================" -ForegroundColor Cyan
        Write-Host "   UNIVERSAL MAIL & SERVER DIAGNOSTIC & REPAIR TOOL       " -ForegroundColor Cyan
        Write-Host "==========================================================" -ForegroundColor Cyan
    }

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
        Write-Host "  1. Test Connectivity & Diagnostics Only (Makes no changes to network)"
        Write-Host "  2. Repair & Test (Changes DNS, Disables IPv6, Tests Connectivity)"
        Write-Host "  3. Rollback (Undo previous network changes)"
        Write-Host "  4. View Network Status (Show physical adapters and current DNS)"
        Write-Host "  5. Inspect Full DNS Records (A, CNAME, MX, SPF, DMARC, NS)"
        Write-Host "  0. Exit Tool" -ForegroundColor Red
        $choice = Read-Host "`nEnter 1, 2, 3, 4, 5, or 0 (Default: 2)"
        
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice -eq "2") { $opMode = "Repair"; $validChoice = $true }
        elseif ($choice -eq "1") { $opMode = "Test"; $validChoice = $true }
        elseif ($choice -eq "3") { $opMode = "Rollback"; $validChoice = $true }
        elseif ($choice -eq "4") { $opMode = "ViewStatus"; $validChoice = $true }
        elseif ($choice -eq "5") { $opMode = "InspectDns"; $validChoice = $true }
        elseif ($choice -eq "0") { try { Stop-Transcript | Out-Null } catch {}; Exit 0 }
        else { Write-Warning "Invalid choice. Please enter a number between 0 and 5.`n" }
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
    $inputHost = Read-Host "Enter the mail server or domain to test (e.g., mail.domain.com or domain.com)"
    if ([string]::IsNullOrWhiteSpace($inputHost)) {
        Write-Warning "No hostname provided."
        if ($isInteractive) {
            Write-Host "`nReturning to Main Menu in 2 seconds..."
            Start-Sleep -Seconds 2
            continue
        } else {
            try { Stop-Transcript | Out-Null } catch {}; Exit 1
        }
    }
    $TargetHosts = $inputHost -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
}

# ---------------------------------------------------------------------------
# EXECUTION: Inspect DNS Mode Only
# ---------------------------------------------------------------------------
if ($opMode -eq "InspectDns") {
    foreach ($target in $TargetHosts) {
        Show-FullDnsRecords -Domain $target
    }
    Write-Host "`nDNS inspection complete." -ForegroundColor Cyan
    if ($isInteractive) {
        $navChoice = Read-Host "`nPress 1 to return to Main Menu, or 0 to Exit [Default: 1]"
        if ($navChoice -eq "0") { Exit 0 }
        continue
    } else {
        try { Stop-Transcript | Out-Null } catch {}; Exit 0
    }
}

# ---------------------------------------------------------------------------
# EXECUTION: View Status Mode
# ---------------------------------------------------------------------------
if ($opMode -eq "ViewStatus") {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    Write-Host "   CURRENT NETWORK & DNS STATUS                           " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan

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
    if ($isInteractive) {
        $navChoice = Read-Host "`nPress 1 to return to Main Menu, or 0 to Exit [Default: 1]"
        if ($navChoice -eq "0") { Exit 0 }
        continue
    } else {
        try { Stop-Transcript | Out-Null } catch {}; Exit 0
    }
}

# ---------------------------------------------------------------------------
# EXECUTION: Rollback Mode
# ---------------------------------------------------------------------------
if ($opMode -eq "Rollback") {
    Write-Host "`n==========================================================" -ForegroundColor Cyan
    if (-not (Test-Path $backupFile)) {
        Write-Error "No backup file found at $backupFile. Nothing to roll back."
        if ($isInteractive) {
            Write-Host "`nReturning to Main Menu in 3 seconds..."
            Start-Sleep -Seconds 3
            continue
        } else {
            try { Stop-Transcript | Out-Null } catch {}; Exit 1
        }
    }

    Write-Host "[+] Rolling back DNS/IPv6 settings from backup..." -ForegroundColor Yellow
    $backup = @(Get-Content $backupFile -Raw | ConvertFrom-Json)

    # Reset global IPv6 prefix policy to Windows defaults
    try {
        netsh interface ipv6 reset prefixpolicy | Out-Null
        Write-Host "[+] Reset global IPv6 prefix policies to default." -ForegroundColor Green
    } catch { }

    foreach ($entry in $backup) {
        # Fallback MAC lookup to handle renamed or re-enabled adapters
        $targetAdapter = Get-NetAdapter | Where-Object { $_.MacAddress -eq $entry.MacAddress } | Select-Object -First 1
        if (-not $targetAdapter) {
            $targetAdapter = Get-NetAdapter -Name $entry.Adapter -ErrorAction SilentlyContinue
        }
        if (-not $targetAdapter) {
            Write-Warning "    [!] Adapter '$($entry.Adapter)' not found or disconnected. Skipping."
            continue
        }
        
        $alias = $targetAdapter.Name
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
                $dnsStr = $entry.OriginalDns -join ", "
                Write-Host "    [OK] Restored original static DNS ($dnsStr) on '$alias'." -ForegroundColor Green
            }
        } catch {
            Write-Warning "    [!] Failed to roll back adapter '$alias' : $_"
        }
    }

    Clear-DnsClientCache
    Write-Host "`nRollback complete." -ForegroundColor Cyan
    if ($isInteractive) {
        $navChoice = Read-Host "`nPress 1 to return to Main Menu, or 0 to Exit [Default: 1]"
        if ($navChoice -eq "0") { Exit 0 }
        continue
    } else {
        try { Stop-Transcript | Out-Null } catch {}; Exit 0
    }
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
        if ($isInteractive) {
            Write-Host "`nReturning to Main Menu in 3 seconds..."
            Start-Sleep -Seconds 3
            continue
        } else {
            try { Stop-Transcript | Out-Null } catch {}; Exit 1
        }
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
            MacAddress     = $adapter.MacAddress
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
            $dnsStr = $selectedDns.IPs -join " / "
            Write-Host "    [OK] $($selectedDns.Name) DNS ($dnsStr) configured." -ForegroundColor Green
        } catch {
            Write-Warning "    [!] Error assigning IPv4 DNS: $_"
        }
    }

    # Prefer IPv4 over IPv6 Globally (Smooth Fallback)
    try {
        netsh interface ipv6 set prefixpolicy ::ffff:0:0/96 46 4 | Out-Null
        Write-Host "`n[+] Global IPv6 prefix policy updated to prefer IPv4." -ForegroundColor Green
    } catch { }

    ConvertTo-Json @($backupEntries) -Depth 3 | Set-Content -Path $backupFile -Encoding UTF8
    Write-Host "`n[+] Pre-change settings backed up to: $backupFile" -ForegroundColor DarkGray
    Write-Host "    Run this script and select 'Rollback' to undo these changes." -ForegroundColor DarkGray

    Write-Host "`n[+] Flushing DNS cache..." -ForegroundColor Yellow
    Clear-DnsClientCache
    Write-Host "    [OK] DNS cache cleared." -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# EXECUTION: Comprehensive Diagnostics
# ---------------------------------------------------------------------------
Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   RUNNING FULL SYSTEM & SERVICE DIAGNOSTICS              " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

$allResults = @{}
$reportText = New-Object System.Text.StringBuilder

$null = $reportText.AppendLine("==========================================================")
$null = $reportText.AppendLine("   IT SUPPORT TOOLS - SERVICE DIAGNOSTIC REPORT           ")
$dateStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $null = $reportText.AppendLine("   Date: $dateStr        ")
$null = $reportText.AppendLine("==========================================================")

foreach ($hostTarget in $TargetHosts) {
    Write-Host "`n----------------------------------------------------------"
    Write-Host "   Testing Target: $hostTarget" -ForegroundColor White
    Write-Host "----------------------------------------------------------"

    $null = $reportText.AppendLine("`nTarget: $hostTarget")

    # If the user inputted a base domain without "mail.", detect MX
    $testedHost = $hostTarget
    if ($hostTarget -notmatch "^(mail|webmail|smtp|imap|pop)\.") {
        try {
            $discoveredMx = Resolve-DnsName -Name $hostTarget -Type MX -ErrorAction SilentlyContinue | Sort-Object Preference | Select-Object -First 1
            if ($discoveredMx -and $discoveredMx.NameExchange) {
                Write-Host "[*] Auto-detected Mail Exchanger (MX): $($discoveredMx.NameExchange)" -ForegroundColor Cyan
                # If target is base domain, we can also test the MX host
            }
        } catch { }
    }

    # Run DNS inspection inline
    Show-FullDnsRecords -Domain $hostTarget

    # Port reachability test suite
    $portsToTest = [ordered]@{
        "DNS_A_Record"       = 0
        "IMAP_993_SSL"       = 993
        "POP3_995_SSL"       = 995
        "SMTP_465_SSL"       = 465
        "SMTP_587_STARTTLS"  = 587
        "Webmail_2096_SSL"   = 2096
        "cPanel_2083_SSL"    = 2083
        "HTTPS_443"          = 443
    }

    $results = [ordered]@{}
    foreach ($p in $portsToTest.Keys) {
        $results[$p] = "NOT TESTED"
    }

    # Test DNS A Record
    $targetIpForSockets = $hostTarget
    Write-Host "`n[*] Verifying Host Resolution..." -ForegroundColor Yellow
    try {
        $dnsResult = Resolve-DnsName -Name $hostTarget -Type A -ErrorAction Stop
        $resolvedIPs = @($dnsResult | Where-Object { $_.IPAddress } | Select-Object -ExpandProperty IPAddress)
        if ($resolvedIPs.Count -gt 0) {
            $ipStr = $resolvedIPs -join ", "
            Write-Host "    [SUCCESS] Host resolved to IP: $ipStr" -ForegroundColor Green
            $results["DNS_A_Record"] = "PASS"
            $targetIpForSockets = $resolvedIPs[0]
        } else {
            Write-Host "    [FAILURE] No IPv4 A records found." -ForegroundColor Red
            $results["DNS_A_Record"] = "FAIL"
        }
    } catch {
        Write-Host "    [FAILURE] Unable to resolve $hostTarget" -ForegroundColor Red
        $results["DNS_A_Record"] = "FAIL"
    }

    # Test Ports using Resolved IP (Prevents Socket DNS Hangs)
    Write-Host "`n[*] Testing Server Ports (Inbound, Outbound, Webmail, Panels)..." -ForegroundColor Yellow
    foreach ($serviceName in $portsToTest.Keys) {
        $portNum = $portsToTest[$serviceName]
        if ($portNum -eq 0) { continue }

        if (Test-PortReachability -HostName $targetIpForSockets -Port $portNum) {
            Write-Host ("    {0,-22} (Port {1,4}) : CONNECTED" -f $serviceName, $portNum) -ForegroundColor Green
            $results[$serviceName] = "PASS"
        } else {
            Write-Host ("    {0,-22} (Port {1,4}) : UNREACHABLE" -f $serviceName, $portNum) -ForegroundColor Red
            $results[$serviceName] = "FAIL"
        }
    }

    # Test SSL Certificate on Port 443 and 993/465
    Write-Host "`n[*] Inspecting SSL/TLS Certificate..." -ForegroundColor Yellow
    $sslCheck = Test-SslCertificate -HostName $hostTarget -Port 443
    if (-not $sslCheck.Valid) {
        # Fallback to test SSL on IMAP 993
        $sslCheck = Test-SslCertificate -HostName $hostTarget -Port 993
    }

    if ($sslCheck.Valid) {
        Write-Host "    [SUCCESS] Certificate is VALID!" -ForegroundColor Green
        Write-Host "    Expires on : $($sslCheck.ExpiresOn) ($($sslCheck.DaysLeft) days remaining)" -ForegroundColor Green
        Write-Host "    Issued To  : $($sslCheck.Subject)" -ForegroundColor DarkGray
        Write-Host "    Issued By  : $($sslCheck.Issuer)" -ForegroundColor DarkGray
        $results["SSL_Certificate"] = "VALID ($($sslCheck.DaysLeft)d left)"
    } else {
        Write-Host "    [WARNING] SSL Check: $($sslCheck.Error)" -ForegroundColor Yellow
        $results["SSL_Certificate"] = "WARNING/FAIL"
    }

    $allResults[$hostTarget] = $results
}

# ---------------------------------------------------------------------------
# 3. Final Summary & Clipboard Export
# ---------------------------------------------------------------------------
Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host "   FINAL SUMMARY ($opMode Mode)                            " -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan

foreach ($hostTarget in $TargetHosts) {
    Write-Host "`n > Results for $hostTarget :" -ForegroundColor White
    $null = $reportText.AppendLine("`n--- Summary for $hostTarget ---")
    
    $hostResults = $allResults[$hostTarget]
    
    foreach ($key in $hostResults.Keys) {
        $status = $hostResults[$key]
        $color = switch -Wildcard ($status) {
            "PASS*"  { "Green" }
            "VALID*" { "Green" }
            "FAIL*"  { "Red" }
            default  { "DarkGray" }
        }
        $line = ("    {0,-22} {1}" -f $key, $status)
        Write-Host $line -ForegroundColor $color
        $null = $reportText.AppendLine($line)
    }
}

Write-Host "`nLog saved to: $logFile" -ForegroundColor DarkGray
Write-Host "Diagnostic complete." -ForegroundColor Cyan

# Option to copy full report to clipboard
Write-Host ""
$copyChoice = Read-Host "Would you like to copy the diagnostic summary to the Clipboard? (Y/N) [Default: Y]"
if ([string]::IsNullOrWhiteSpace($copyChoice) -or $copyChoice -match "^[Yy]") {
    try {
        Set-Clipboard -Value $reportText.ToString()
        Write-Host "[OK] Full report successfully copied to clipboard! You can now paste (Ctrl+V) into a ticket or chat." -ForegroundColor Green
    } catch {
        Write-Warning "Could not access clipboard in this session."
    }
}

if ($isInteractive) {
    $navChoice = Read-Host "`nPress 1 to return to Main Menu, or 0 to Exit [Default: 1]"
    if ($navChoice -eq "0") { Exit 0 }
    continue
} else {
    try { Stop-Transcript | Out-Null } catch {}; Exit 0
}
}





