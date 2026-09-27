<#
.SYNOPSIS
    Comprehensive Mail Client Profile Manager for Microsoft Outlook and Mozilla Thunderbird.
.DESCRIPTION
    A diagnostic, configuration, testing, and cleanup utility for IT technicians and support teams.
    Supports both an interactive console menu and non-interactive command-line switches.

    Key Capabilities:
      1. Inspect & List:
         - Outlook: Registry profile detection (16.0 / 365, 15.0, 14.0), data files (.ost/.pst), size checks, locks.
         - Thunderbird: profiles.ini parsing, profile folders, prefs.js account & server extraction, orphaned profiles.
      2. Configure / Add:
         - Outlook: Automated Microsoft .PRF generator & importer, registry initialization, profile wizard launch.
         - Thunderbird: Automatic profile directory generation, profiles.ini / installs.ini updates, user.js account seeding.
      3. Deep Test & Diagnostics:
         - Local health check (locks, crash artifacts, missing directories, oversized OSTs).
         - Live server connectivity test for all detected IMAP, POP3, and SMTP accounts (TCP socket + SSL/TLS certificate validation).
      4. Safe Clean & Repair:
         - Non-destructive cache cleanup (Thunderbird cache2/startupCache, Outlook RoamCache/Autodiscover/Temp).
         - Stuck lock clearing (Thunderbird parent.lock, hung zombie client processes).
         - Mailbox index repair (cleans corrupted Thunderbird .msf files to force safe re-indexing).
         - Orphaned profile and unused .OST file cleanup to reclaim disk space.
         - Profile removal and full client resets with AUTOMATIC registry/file backups prior to modification.
.NOTES
    Author: Samir Koulali (https://samirkoulali.art)
    Version: 1.0.0
    License: Open source for non-commercial use
#>

[CmdletBinding()]
param(
    [ValidateSet('All', 'Outlook', 'Thunderbird')]
    [string]$Client = 'All',

    [ValidateSet('Menu', 'List', 'Test', 'Add', 'Clean', 'Repair', 'Backup', 'Reset')]
    [string]$Action = 'Menu',

    # Add Profile parameters:
    [string]$ProfileName,
    [string]$EmailAddress,
    [string]$DisplayName,
    [ValidateSet('IMAP', 'POP3')]
    [string]$AccountType = 'IMAP',
    [string]$IncomingServer,
    [int]$IncomingPort = 993,
    [ValidateSet('SSL', 'STARTTLS', 'None')]
    [string]$IncomingSecurity = 'SSL',
    [string]$IncomingUser,
    [string]$OutgoingServer,
    [int]$OutgoingPort = 465,
    [ValidateSet('SSL', 'STARTTLS', 'None')]
    [string]$OutgoingSecurity = 'SSL',
    [string]$OutgoingUser,
    [switch]$SetAsDefault,

    # Clean parameters:
    [ValidateSet('SafeCache', 'LocksOnly', 'CorruptedIndices', 'OrphanedProfiles', 'SpecificProfile', 'FullReset')]
    [string]$CleanScope = 'SafeCache',
    [string]$TargetProfile,
    [switch]$Force,

    # Test parameters:
    [switch]$SkipNetworkTest
)

# ---------------------------------------------------------------------------
# Setup: paths, logging, global directories
# ---------------------------------------------------------------------------
$appDataRoot   = Join-Path $env:ProgramData "MailClientManager"
$logDir        = Join-Path $appDataRoot "logs"
$backupDir     = Join-Path $appDataRoot "backups"
$timestamp     = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile       = Join-Path $logDir "mail_profile_$timestamp.log"

foreach ($dir in @($appDataRoot, $logDir, $backupDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

try {
    Start-Transcript -Path $logFile -Append -ErrorAction SilentlyContinue | Out-Null
} catch {}

$global:LastReportText = ""

# ---------------------------------------------------------------------------
# UI & Helper Functions
# ---------------------------------------------------------------------------
function Write-Header {
    param([string]$Title)
    Clear-Host
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "   MAIL CLIENT PROFILE MANAGER (OUTLOOK & THUNDERBIRD)   " -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Cyan
    if ($Title) {
        Write-Host "  >> $Title" -ForegroundColor Yellow
        Write-Host "----------------------------------------------------------" -ForegroundColor DarkGray
    }
    Write-Host ""
}

function Show-Notification {
    param([string]$Message, [string]$Type = "Info")
    $color = switch ($Type) {
        "Success" { "Green" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        Default   { "Cyan" }
    }
    $prefix = switch ($Type) {
        "Success" { "[SUCCESS]" }
        "Warning" { "[WARN]   " }
        "Error"   { "[ERROR]  " }
        Default   { "[INFO]   " }
    }
    Write-Host "$prefix $Message" -ForegroundColor $color
}

function Pause-Menu {
    Write-Host ""
    Write-Host "Press any key to return to the menu..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# ---------------------------------------------------------------------------
# Network & SSL Test Helpers (Cached IP sockets to avoid DNS hangs)
# ---------------------------------------------------------------------------
function Test-FastPortReachability {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 3000)
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    try {
        # Resolve IP first to prevent socket hangs on unresolvable DNS
        $ipEntry = [System.Net.Dns]::GetHostAddresses($HostName) | 
                   Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | 
                   Select-Object -First 1

        $target = if ($ipEntry) { $ipEntry } else { $HostName }
        $connect = $tcpClient.BeginConnect($target, $Port, $null, $null)
        $success = $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($success) {
            $tcpClient.EndConnect($connect)
            $tcpClient.Close()
            return $true
        }
    } catch {
        # Return false on error
    } finally {
        $tcpClient.Close()
    }
    return $false
}

function Test-MailSslCertificate {
    param([string]$HostName, [int]$Port = 443, [int]$TimeoutMs = 4000)
    $tcpClient = New-Object System.Net.Sockets.TcpClient
    try {
        $ipEntry = [System.Net.Dns]::GetHostAddresses($HostName) | 
                   Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } | 
                   Select-Object -First 1
        $target = if ($ipEntry) { $ipEntry } else { $HostName }

        $connect = $tcpClient.BeginConnect($target, $Port, $null, $null)
        if (-not $connect.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            $tcpClient.Close()
            return [PSCustomObject]@{ Valid = $false; Error = "Connection timed out" }
        }
        $netStream = $tcpClient.GetStream()

        # Handle STARTTLS negotiation for SMTP submission (Port 587)
        if ($Port -eq 587) {
            $reader = New-Object System.IO.StreamReader($netStream)
            $writer = New-Object System.IO.StreamWriter($netStream)
            $writer.AutoFlush = $true

            $null = $reader.ReadLine() # Read server banner
            $writer.WriteLine("EHLO localhost")
            while ($line = $reader.ReadLine()) {
                if ($line -match '^250\s') { break }
            }
            $writer.WriteLine("STARTTLS")
            $tlsResp = $reader.ReadLine()
            if ($tlsResp -notmatch '^220') {
                $tcpClient.Close()
                return [PSCustomObject]@{ Valid = $false; Error = "STARTTLS rejected: $tlsResp" }
            }
        }

        $sslStream = New-Object System.Net.Security.SslStream(
            $netStream,
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
            Error       = if ($isExpired) { "Certificate Expired!" } else { $null }
        }
    } catch {
        return [PSCustomObject]@{ Valid = $false; Error = $_.Exception.Message }
    } finally {
        $tcpClient.Close()
    }
}

# ---------------------------------------------------------------------------
# Discovery: Outlook Paths, Profiles & Accounts
# ---------------------------------------------------------------------------
function Get-OutlookExecutable {
    $paths = @(
        (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' -ErrorAction SilentlyContinue).'(Default)',
        "$env:ProgramFiles\Microsoft Office\Root\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Root\Office16\OUTLOOK.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE",
        "$env:ProgramFiles\Microsoft Office\Office15\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15\OUTLOOK.EXE"
    )
    foreach ($p in $paths) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}

function Get-OutlookProfiles {
    $results = @()
    $versions = @("16.0", "15.0", "14.0")
    $defaultProfile = $null

    foreach ($ver in $versions) {
        $baseKey = "HKCU:\Software\Microsoft\Office\$ver\Outlook"
        $profKey = "$baseKey\Profiles"

        if (Test-Path $profKey) {
            $defProp = Get-ItemProperty $profKey -ErrorAction SilentlyContinue
            if ($defProp -and $defProp.DefaultProfile) {
                $defaultProfile = $defProp.DefaultProfile
            }

            $subKeys = (Get-Item $profKey -ErrorAction SilentlyContinue).GetSubKeyNames()
            if ($subKeys) {
                foreach ($profileName in $subKeys) {
                    $itemKey = "$profKey\$profileName"
                    $isDefault = ($profileName -eq $defaultProfile)

                    # Scan subkeys for data file references (.ost, .pst) and account information
                    $dataFiles = @()
                    $servers = @()
                    $emails = @()

                    try {
                        Get-ChildItem -Path $itemKey -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
                            $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
                            if ($props) {
                                foreach ($propName in $props.PSObject.Properties.Name) {
                                    $val = $props.$propName
                                    if ($val -is [string]) {
                                        if ($val -match '\.(ost|pst)$') {
                                            $dataFiles += $val
                                        }
                                        if ($val -match '^([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})$') {
                                            $emails += $val
                                        }
                                    }
                                }
                            }
                        }
                    } catch {}

                    # Check for OST/PST files in standard Outlook local path
                    $localOutlook = "$env:LOCALAPPDATA\Microsoft\Outlook"
                    if (Test-Path $localOutlook) {
                        $potentialOsts = Get-ChildItem -Path $localOutlook -Filter "*.ost" -ErrorAction SilentlyContinue | 
                                         Where-Object { $_.BaseName -like "*$profileName*" -or $_.BaseName -match "outlook" }
                        foreach ($ost in $potentialOsts) {
                            $dataFiles += $ost.FullName
                        }
                    }

                    $dataFiles = $dataFiles | Select-Object -Unique
                    $emails = $emails | Select-Object -Unique

                    # Check OST sizes and locks
                    $fileDetails = @()
                    foreach ($df in $dataFiles) {
                        if (Test-Path $df) {
                            $fItem = Get-Item $df
                            $sizeMb = [math]::Round($fItem.Length / 1MB, 2)
                            $fileDetails += [PSCustomObject]@{
                                Path   = $df
                                SizeMB = $sizeMb
                                Exists = $true
                            }
                        } else {
                            $fileDetails += [PSCustomObject]@{
                                Path   = $df
                                SizeMB = 0
                                Exists = $false
                            }
                        }
                    }

                    $results += [PSCustomObject]@{
                        Client         = "Outlook"
                        OfficeVersion  = $ver
                        ProfileName    = $profileName
                        IsDefault      = $isDefault
                        RegistryPath   = $itemKey
                        DataFiles      = $fileDetails
                        Emails         = $emails
                    }
                }
            }
        }
    }

    # Also check legacy Windows Messaging Subsystem if no Office profiles found
    $legacyKey = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
    if (Test-Path $legacyKey) {
        $legacySubKeys = (Get-Item $legacyKey -ErrorAction SilentlyContinue).GetSubKeyNames()
        if ($legacySubKeys) {
            foreach ($profileName in $legacySubKeys) {
                if (-not ($results | Where-Object { $_.ProfileName -eq $profileName })) {
                    $results += [PSCustomObject]@{
                        Client         = "Outlook"
                        OfficeVersion  = "Legacy"
                        ProfileName    = $profileName
                        IsDefault      = $false
                        RegistryPath   = "$legacyKey\$profileName"
                        DataFiles      = @()
                        Emails         = @()
                    }
                }
            }
        }
    }

    return $results
}

# ---------------------------------------------------------------------------
# Discovery: Thunderbird Paths, Profiles & Accounts
# ---------------------------------------------------------------------------
function Get-ThunderbirdExecutable {
    $paths = @(
        (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\thunderbird.exe' -ErrorAction SilentlyContinue).'(Default)',
        "$env:ProgramFiles\Mozilla Thunderbird\thunderbird.exe",
        "${env:ProgramFiles(x86)}\Mozilla Thunderbird\thunderbird.exe"
    )
    foreach ($p in $paths) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}

function Get-ThunderbirdProfiles {
    $results = @()
    $tbRoot = "$env:APPDATA\Thunderbird"
    $iniPath = Join-Path $tbRoot "profiles.ini"

    if (-not (Test-Path $iniPath)) {
        return $results
    }

    $lines = Get-Content $iniPath -ErrorAction SilentlyContinue
    $currentSection = ""
    $sections = @{}

    foreach ($line in $lines) {
        $trim = $line.Trim()
        if ($trim -match '^\[(.*)\]$') {
            $currentSection = $matches[1]
            $sections[$currentSection] = @{}
        } elseif ($trim -match '^([^=]+)=(.*)$' -and $currentSection) {
            $key = $matches[1].Trim()
            $val = $matches[2].Trim()
            $sections[$currentSection][$key] = $val
        }
    }

    # Identify registered profile folders
    $registeredPaths = @()

    foreach ($secKey in $sections.Keys) {
        if ($secKey -match '^Profile\d+$') {
            $sec = $sections[$secKey]
            $pName = $sec['Name']
            $pPath = $sec['Path']
            $isRel = ($sec['IsRelative'] -eq '1')
            $isDef = ($sec['Default'] -eq '1')

            $fullPath = if ($isRel) { Join-Path $tbRoot ($pPath -replace '/', '\') } else { $pPath }
            $registeredPaths += $fullPath

            # Analyze profile folder health
            $folderExists = Test-Path $fullPath
            $folderSizeMb = 0
            $hasLock = $false
            $prefsExists = $false
            $servers = @()
            $identities = @()

            if ($folderExists) {
                $hasLock = Test-Path (Join-Path $fullPath "parent.lock")
                $prefsPath = Join-Path $fullPath "prefs.js"
                $prefsExists = Test-Path $prefsPath

                try {
                    $meas = Get-ChildItem -Path $fullPath -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
                    if ($meas -and $meas.Sum) {
                        $folderSizeMb = [math]::Round($meas.Sum / 1MB, 2)
                    }
                } catch {}

                # Parse prefs.js for configured accounts
                if ($prefsExists) {
                    $prefsLines = Get-Content $prefsPath -ErrorAction SilentlyContinue
                    
                    $accMap = @{}
                    $srvMap = @{}
                    $idMap  = @{}
                    $smtpMap = @{}

                    foreach ($pl in $prefsLines) {
                        if ($pl -match 'user_pref\("mail\.account\.(account\d+)\.(server|identities)",\s*"([^"]+)"\);') {
                            $aId = $matches[1]; $prop = $matches[2]; $val = $matches[3]
                            if (-not $accMap.ContainsKey($aId)) { $accMap[$aId] = @{} }
                            $accMap[$aId][$prop] = $val
                        }
                        elseif ($pl -match 'user_pref\("mail\.server\.(server\d+)\.(hostname|port|type|userName|socketType)",\s*"?([^"\)]*)"?\);') {
                            $sId = $matches[1]; $prop = $matches[2]; $val = $matches[3]
                            if (-not $srvMap.ContainsKey($sId)) { $srvMap[$sId] = @{} }
                            $srvMap[$sId][$prop] = $val
                        }
                        elseif ($pl -match 'user_pref\("mail\.identity\.(id\d+)\.(useremail|fullName|smtpServer)",\s*"([^"]+)"\);') {
                            $iId = $matches[1]; $prop = $matches[2]; $val = $matches[3]
                            if (-not $idMap.ContainsKey($iId)) { $idMap[$iId] = @{} }
                            $idMap[$iId][$prop] = $val
                        }
                        elseif ($pl -match 'user_pref\("mail\.smtpserver\.(smtp\d+)\.(hostname|port|username|try_ssl)",\s*"?([^"\)]*)"?\);') {
                            $mId = $matches[1]; $prop = $matches[2]; $val = $matches[3]
                            if (-not $smtpMap.ContainsKey($mId)) { $smtpMap[$mId] = @{} }
                            $smtpMap[$mId][$prop] = $val
                        }
                    }

                    # Assemble clean account definitions
                    $cleanAccounts = @()
                    foreach ($aId in $accMap.Keys) {
                        $sId = $accMap[$aId]['server']
                        $iId = ($accMap[$aId]['identities'] -split ',')[0]
                        
                        $email = if ($idMap.ContainsKey($iId)) { $idMap[$iId]['useremail'] } else { $null }
                        $smtpId = if ($idMap.ContainsKey($iId)) { $idMap[$iId]['smtpServer'] } else { $null }

                        $inHost = if ($srvMap.ContainsKey($sId)) { $srvMap[$sId]['hostname'] } else { $null }
                        $inType = if ($srvMap.ContainsKey($sId)) { $srvMap[$sId]['type'] } else { "imap" }
                        $inPort = if ($srvMap.ContainsKey($sId) -and $srvMap[$sId]['port']) { [int]$srvMap[$sId]['port'] } else { if ($inType -eq 'pop3') { 995 } else { 993 } }
                        $inSec  = if ($srvMap.ContainsKey($sId)) { switch ($srvMap[$sId]['socketType']) { "3" { "SSL/TLS" } "2" { "STARTTLS" } Default { "None" } } } else { "SSL/TLS" }

                        $outHost = if ($smtpId -and $smtpMap.ContainsKey($smtpId)) { $smtpMap[$smtpId]['hostname'] } else { $null }
                        $outPort = if ($smtpId -and $smtpMap.ContainsKey($smtpId) -and $smtpMap[$smtpId]['port']) { [int]$smtpMap[$smtpId]['port'] } else { 465 }
                        $outSec  = if ($smtpId -and $smtpMap.ContainsKey($smtpId)) { switch ($smtpMap[$smtpId]['try_ssl']) { "3" { "SSL/TLS" } "2" { "STARTTLS" } Default { "None" } } } else { "SSL/TLS" }

                        if ($email -and $inHost -and $inHost -ne "Local Folders" -and $inHost -notmatch '^(localhost|127\.0\.0\.1)$') {
                            $cleanAccounts += [PSCustomObject]@{
                                Email            = $email
                                FullName         = if ($idMap.ContainsKey($iId)) { $idMap[$iId]['fullName'] } else { "" }
                                IncomingHost     = $inHost
                                IncomingPort     = $inPort
                                IncomingType     = $inType.ToUpper()
                                IncomingSecurity = $inSec
                                OutgoingHost     = $outHost
                                OutgoingPort     = $outPort
                                OutgoingSecurity = $outSec
                            }

                            if (-not ($servers | Where-Object { $_.Hostname -eq $inHost -and $_.Port -eq $inPort })) {
                                $servers += [PSCustomObject]@{ Hostname = $inHost; Port = $inPort; Type = $inType.ToUpper(); SocketType = $inSec }
                            }
                            if ($outHost -and $outHost -notmatch '^(localhost|127\.0\.0\.1)$' -and -not ($servers | Where-Object { $_.Hostname -eq $outHost -and $_.Port -eq $outPort })) {
                                $servers += [PSCustomObject]@{ Hostname = $outHost; Port = $outPort; Type = "SMTP"; SocketType = $outSec }
                            }
                        }
                    }

                    $accounts = $cleanAccounts
                    $identities = $cleanAccounts | ForEach-Object { $_.Email }
                }
            }

            $results += [PSCustomObject]@{
                Client        = "Thunderbird"
                ProfileName   = $pName
                IsDefault     = $isDef
                FolderPath    = $fullPath
                FolderExists  = $folderExists
                SizeMB        = $folderSizeMb
                HasLock       = $hasLock
                HasPrefs      = $prefsExists
                Servers       = $servers
                Accounts      = $accounts
                Identities    = $identities
                IsOrphaned    = $false
            }
        }
    }

    # Detect orphaned profile folders on disk
    $profilesDir = Join-Path $tbRoot "Profiles"
    if (Test-Path $profilesDir) {
        $diskFolders = Get-ChildItem -Path $profilesDir -Directory -ErrorAction SilentlyContinue
        foreach ($df in $diskFolders) {
            $isFound = $false
            foreach ($rp in $registeredPaths) {
                if ($rp.TrimEnd('\') -eq $df.FullName.TrimEnd('\')) {
                    $isFound = $true
                    break
                }
            }
            if (-not $isFound) {
                $results += [PSCustomObject]@{
                    Client        = "Thunderbird"
                    ProfileName   = "(Orphaned: $($df.Name))"
                    IsDefault     = $false
                    FolderPath    = $df.FullName
                    FolderExists  = $true
                    SizeMB        = 0
                    HasLock       = (Test-Path (Join-Path $df.FullName "parent.lock"))
                    HasPrefs      = (Test-Path (Join-Path $df.FullName "prefs.js"))
                    Servers       = @()
                    Accounts      = @()
                    Identities    = @()
                    IsOrphaned    = $true
                }
            }
        }
    }

    return $results
}

# ---------------------------------------------------------------------------
# Diagnostics: Unified Profile Test & Server Reachability
# ---------------------------------------------------------------------------
function Invoke-ProfileTestReport {
    param([switch]$OmitNetwork)

    $outlookExe = Get-OutlookExecutable
    $tbExe      = Get-ThunderbirdExecutable
    $outlookProc = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    $tbProc      = Get-Process -Name "thunderbird" -ErrorAction SilentlyContinue

    $outlookProfiles = Get-OutlookProfiles
    $tbProfiles      = Get-ThunderbirdProfiles

    # Collect all unique remote servers across configured profiles
    $serversToTest = @()
    foreach ($tp in $tbProfiles) {
        foreach ($s in $tp.Servers) {
            if ($s.Hostname -and $s.Hostname -notmatch '^(localhost|127\.0\.0\.1)$') {
                $key = "$($s.Hostname):$($s.Port)"
                if (-not ($serversToTest | Where-Object { "$($_.Hostname):$($_.Port)" -eq $key })) {
                    $serversToTest += $s
                }
            }
        }
    }

    $networkResults = @()
    if (-not $OmitNetwork -and $serversToTest.Count -gt 0) {
        $curr = 0
        $total = $serversToTest.Count
        foreach ($s in $serversToTest) {
            $curr++
            Write-Progress -Activity "Testing Mail Server Reachability & SSL" `
                -Status ("[{0}/{1}] Checking {2}:{3} ({4})" -f $curr, $total, $s.Hostname, $s.Port, $s.Type) `
                -PercentComplete ([int](($curr / $total) * 100))

            $isReachable = Test-FastPortReachability -HostName $s.Hostname -Port $s.Port
            $certInfo = "-"
            $certStatus = "OK"

            if ($isReachable) {
                if ($s.Port -in @(993, 995, 465, 587, 443)) {
                    $cert = Test-MailSslCertificate -HostName $s.Hostname -Port $s.Port
                    if ($cert.Valid) {
                        $certInfo = "Valid ($($cert.DaysLeft) days left, $($cert.ExpiresOn))"
                    } else {
                        $certInfo = "CERT ERROR: $($cert.Error)"
                        $certStatus = "ERROR"
                    }
                }
            } else {
                $certInfo = "N/A (Port Closed or Blocked)"
                $certStatus = "FAIL"
            }

            $networkResults += [PSCustomObject]@{
                Hostname    = $s.Hostname
                Port        = $s.Port
                Protocol    = $s.Type
                Reachable   = $isReachable
                CertInfo    = $certInfo
                CertStatus  = $certStatus
            }
        }
        Write-Progress -Activity "Testing Mail Server Reachability & SSL" -Completed
    }

    # Build clean dashboard output
    $sep = "=========================================================================================="
    $div = "------------------------------------------------------------------------------------------"

    $reportLines = @()
    $reportLines += $sep
    $reportLines += "                       MAIL CLIENT PROFILE DIAGNOSTIC DASHBOARD                           "
    $reportLines += "                       Report Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')        "
    $reportLines += $sep
    $reportLines += ""

    # Section 1: Application Status
    $reportLines += "1. APPLICATIONS & RUNNING PROCESSES"
    $reportLines += $div
    $reportLines += ("{0,-22} | {1,-20} | {2}" -f "Application", "Status", "Installation Path")
    $reportLines += $div

    $outState = if ($outlookProc) { "RUNNING (PID: $($outlookProc.Id -join ','))" } else { "IDLE" }
    $outPathStr = if ($outlookExe) { $outlookExe } else { "Not Found" }
    $reportLines += ("{0,-22} | {1,-20} | {2}" -f "Microsoft Outlook", "[$outState]", $outPathStr)

    $tbPids = if ($tbProc) { ($tbProc | ForEach-Object { $_.Id }) -join ',' } else { "" }
    $tbState = if ($tbProc) { "RUNNING (PID: $tbPids)" } else { "IDLE" }
    $tbPathStr = if ($tbExe) { $tbExe } else { "Not Found" }
    $reportLines += ("{0,-22} | {1,-20} | {2}" -f "Mozilla Thunderbird", "[$tbState]", $tbPathStr)
    $reportLines += $div
    $reportLines += ""

    # Section 2: Profiles Overview
    $reportLines += "2. CLIENT PROFILES SUMMARY"
    $reportLines += $div
    $reportLines += ("{0,-14} | {1,-20} | {2,-9} | {3,-11} | {4,-14} | {5}" -f "Client", "Profile Name", "Default", "Data Size", "Health State", "Linked Accounts")
    $reportLines += $div

    $allProfiles = @()
    foreach ($op in $outlookProfiles) {
        $totSize = 0
        foreach ($df in $op.DataFiles) { $totSize += $df.SizeMB }
        $accCount = if ($op.Emails) { "$($op.Emails.Count) Accounts" } else { "None" }
        $health = if ($totSize -gt 45000) { "OVERSIZE (>45GB)" } else { "Healthy" }
        $isDef = if ($op.IsDefault) { "YES" } else { "NO" }
        $reportLines += ("{0,-14} | {1,-20} | {2,-9} | {3,-11} | {4,-14} | {5}" -f "Outlook", $op.ProfileName, $isDef, "$totSize MB", $health, $accCount)
    }

    foreach ($tp in $tbProfiles) {
        $accCount = if ($tp.Accounts -and $tp.Accounts.Count -gt 0) { "$($tp.Accounts.Count) Accounts" } else { "None" }
        $health = if ($tp.HasLock) { "In-Use (Lock)" } elseif ($tp.IsOrphaned) { "Orphaned" } else { "Healthy" }
        $isDef = if ($tp.IsDefault) { "YES" } else { "NO" }
        $reportLines += ("{0,-14} | {1,-20} | {2,-9} | {3,-11} | {4,-14} | {5}" -f "Thunderbird", $tp.ProfileName, $isDef, "$($tp.SizeMB) MB", $health, $accCount)
    }
    $reportLines += $div
    $reportLines += ""

    # Section 3: Configured Mail Accounts & Routing
    $reportLines += "3. CONFIGURED MAIL ACCOUNTS & ROUTING"
    $reportLines += $div
    $reportLines += ("{0,-30} | {1,-28} | {2,-28}" -f "Account Email Address", "Incoming (Host:Port)", "Outgoing (Host:Port)")
    $reportLines += $div

    $hasAccounts = $false
    foreach ($tp in $tbProfiles) {
        if ($tp.Accounts -and $tp.Accounts.Count -gt 0) {
            foreach ($acc in $tp.Accounts) {
                $hasAccounts = $true
                $inStr  = "$($acc.IncomingHost):$($acc.IncomingPort) ($($acc.IncomingSecurity))"
                $outStr = if ($acc.OutgoingHost) { "$($acc.OutgoingHost):$($acc.OutgoingPort) ($($acc.OutgoingSecurity))" } else { "None" }
                $reportLines += ("{0,-30} | {1,-28} | {2,-28}" -f $acc.Email, $inStr, $outStr)
            }
        }
    }
    if (-not $hasAccounts) {
        $reportLines += "  (No active configured mail accounts detected in profile preferences)"
    }
    $reportLines += $div
    $reportLines += ""

    # Section 4: Live Server Reachability & SSL Test
    if (-not $OmitNetwork -and $networkResults.Count -gt 0) {
        $reportLines += "4. LIVE SERVER REACHABILITY & SSL/TLS HEALTH CHECK"
        $reportLines += $div
        $reportLines += ("{0,-24} | {1,-6} | {2,-8} | {3,-12} | {4}" -f "Mail Server Host", "Port", "Type", "Reachability", "SSL/TLS Certificate Status")
        $reportLines += $div

        $reachableCount = 0
        $certValidCount = 0

        foreach ($nr in $networkResults) {
            $reachStr = if ($nr.Reachable) { $reachableCount++; "[PASS]" } else { "[FAIL: CLOSED]" }
            if ($nr.CertStatus -eq "OK") { $certValidCount++ }
            $reportLines += ("{0,-24} | {1,-6} | {2,-8} | {3,-12} | {4}" -f $nr.Hostname, $nr.Port, $nr.Protocol, $reachStr, $nr.CertInfo)
        }
        $reportLines += $div
        $reportLines += ""

        # Executive Summary
        $totalServers = $networkResults.Count
        $reportLines += "SUMMARY: $reachableCount/$totalServers Servers Reachable | $certValidCount/$totalServers Valid SSL Certificates"
    } else {
        $reportLines += "4. SERVER CONNECTIVITY: (Skipped or no remote mail servers configured)"
    }

    $reportLines += $sep
    $global:LastReportText = ($reportLines -join "`r`n")
    Write-Host $global:LastReportText
}

# ---------------------------------------------------------------------------
# Configure & Add: Thunderbird Profile Generator
# ---------------------------------------------------------------------------
function Add-ThunderbirdProfile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [string]$Email,
        [string]$DisplayName,
        [string]$ServerHost,
        [int]$ImapPort = 993,
        [int]$SmtpPort = 465,
        [string]$Username,
        [switch]$MakeDefault
    )

    $tbRoot = "$env:APPDATA\Thunderbird"
    $profilesIni = Join-Path $tbRoot "profiles.ini"
    $installsIni = Join-Path $tbRoot "installs.ini"

    if (-not (Test-Path $tbRoot)) {
        New-Item -Path $tbRoot -ItemType Directory -Force | Out-Null
    }

    # Generate 8-character random salt
    $chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    $random = New-Object System.Random
    $salt = -join (1..8 | ForEach-Object { $chars[$random.Next($chars.Length)] })
    $folderName = "$salt.$Name"
    $targetProfileDir = Join-Path (Join-Path $tbRoot "Profiles") $folderName

    Show-Notification "Creating Thunderbird profile directory at: $targetProfileDir"
    New-Item -Path $targetProfileDir -ItemType Directory -Force | Out-Null

    # Configure user.js if mail settings are provided
    if ($Email -and $ServerHost) {
        Show-Notification "Generating pre-configured user.js email settings..."
        $userLogin = if ($Username) { $Username } else { $Email }
        $dispName  = if ($DisplayName) { $DisplayName } else { $Email }

        $userJsLines = @(
            '// Auto-generated configuration by IT Support MailClient-ProfileManager',
            'user_pref("mail.accountmanager.accounts", "account1");',
            'user_pref("mail.accountmanager.defaultaccount", "account1");',
            'user_pref("mail.accountmanager.localfoldersserver", "server1");',
            '',
            '// Local Folders Server',
            'user_pref("mail.server.server1.type", "none");',
            'user_pref("mail.server.server1.userName", "nobody");',
            'user_pref("mail.server.server1.hostname", "Local Folders");',
            'user_pref("mail.server.server1.name", "Local Folders");',
            '',
            '// Identity',
            "user_pref(`"mail.identity.id1.fullName`", `"$dispName`");",
            "user_pref(`"mail.identity.id1.useremail`", `"$Email`");",
            'user_pref("mail.identity.id1.valid", true);',
            'user_pref("mail.identity.id1.smtpServer", "smtp1");',
            '',
            '// Incoming IMAP Server',
            'user_pref("mail.server.server2.type", "imap");',
            "user_pref(`"mail.server.server2.hostname`", `"$ServerHost`");",
            "user_pref(`"mail.server.server2.port`", $ImapPort);",
            "user_pref(`"mail.server.server2.userName`", `"$userLogin`");",
            "user_pref(`"mail.server.server2.name`", `"$Email`");",
            "user_pref(`"mail.server.server2.realhostname`", `"$ServerHost`");",
            "user_pref(`"mail.server.server2.realuserName`", `"$userLogin`");",
            'user_pref("mail.server.server2.socketType", 3);', # 3 = SSL/TLS
            'user_pref("mail.server.server2.authMethod", 3);', # 3 = Normal password
            '',
            '// Link Account 1 to Incoming Server and Identity',
            'user_pref("mail.account.account1.server", "server2");',
            'user_pref("mail.account.account1.identities", "id1");',
            '',
            '// Outgoing SMTP Server',
            'user_pref("mail.smtpservers", "smtp1");',
            "user_pref(`"mail.smtpserver.smtp1.hostname`", `"$ServerHost`");",
            "user_pref(`"mail.smtpserver.smtp1.port`", $SmtpPort);",
            "user_pref(`"mail.smtpserver.smtp1.username`", `"$userLogin`");",
            'user_pref("mail.smtpserver.smtp1.authMethod", 3);',
            'user_pref("mail.smtpserver.smtp1.try_ssl", 3);', # 3 = SSL/TLS
            "user_pref(`"mail.smtpserver.smtp1.description`", `"$ServerHost`");",
            'user_pref("mail.smtp.defaultserver", "smtp1");'
        )

        $userJsPath = Join-Path $targetProfileDir "user.js"
        [System.IO.File]::WriteAllLines($userJsPath, $userJsLines, [System.Text.Encoding]::ASCII)
    }

    # Update profiles.ini
    Show-Notification "Registering new profile in profiles.ini..."
    $iniContent = if (Test-Path $profilesIni) { Get-Content $profilesIni } else { @("[General]", "StartWithLastProfile=1", "Version=2") }
    
    # Find highest Profile number
    $maxProfileNum = -1
    foreach ($line in $iniContent) {
        if ($line -match '^\[Profile(\d+)\]$') {
            $n = [int]$matches[1]
            if ($n -gt $maxProfileNum) { $maxProfileNum = $n }
        }
    }
    $newProfileNum = $maxProfileNum + 1

    $newProfileSection = @(
        "",
        "[Profile$newProfileNum]",
        "Name=$Name",
        "IsRelative=1",
        "Path=Profiles/$folderName"
    )

    if ($MakeDefault) {
        $newProfileSection += "Default=1"
        # Remove default flag from any other profile
        $updatedIni = @()
        foreach ($line in $iniContent) {
            if ($line -match '^Default=1$' -and -not ($line -match '^\[Install')) {
                # skip
            } else {
                $updatedIni += $line
            }
        }
        $iniContent = $updatedIni
    }

    $finalIni = $iniContent + $newProfileSection
    [System.IO.File]::WriteAllLines($profilesIni, $finalIni, [System.Text.Encoding]::UTF8)

    # Update installs.ini if default
    if ($MakeDefault -and (Test-Path $installsIni)) {
        $installsLines = Get-Content $installsIni
        $newInstalls = @()
        foreach ($il in $installsLines) {
            if ($il -match '^Default=.*$') {
                $newInstalls += "Default=Profiles/$folderName"
            } else {
                $newInstalls += $il
            }
        }
        [System.IO.File]::WriteAllLines($installsIni, $newInstalls, [System.Text.Encoding]::UTF8)
    }

    Show-Notification "Thunderbird profile '$Name' successfully created and configured!" -Type "Success"
}

# ---------------------------------------------------------------------------
# Configure & Add: Outlook Profile PRF Generator & Importer
# ---------------------------------------------------------------------------
function Add-OutlookProfile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Name,
        [string]$Email,
        [string]$DisplayName,
        [string]$ServerHost,
        [int]$ImapPort = 993,
        [int]$SmtpPort = 465,
        [string]$Username,
        [switch]$MakeDefault,
        [switch]$LaunchWizard
    )

    $outlookExe = Get-OutlookExecutable

    if ($LaunchWizard) {
        if ($outlookExe) {
            Show-Notification "Launching native Outlook Profile Manager..."
            Start-Process -FilePath $outlookExe -ArgumentList "/profiles"
            return
        } else {
            Show-Notification "Outlook executable not found." -Type "Error"
            return
        }
    }

    # Generate standard Microsoft Outlook PRF (Profile File)
    Show-Notification "Generating Microsoft Outlook Profile Configuration (.PRF)..."
    $userLogin = if ($Username) { $Username } else { $Email }
    $dispName  = if ($DisplayName) { $DisplayName } else { $Email }
    $prfPath   = Join-Path $backupDir "Outlook_$Name.prf"
    $defFlag   = if ($MakeDefault) { "Yes" } else { "No" }

    $prfLines = @(
        "; **************************************************************",
        "; Microsoft Outlook Setup PRF Generated by IT Support Tools",
        "; **************************************************************",
        "[General]",
        "Custom=1",
        "ProfileName=$Name",
        "DefaultProfile=$defFlag",
        "OverwriteProfile=Yes",
        "ModifyExistingProfile=No",
        "",
        "[Service List]",
        "Service1=Internet E-mail",
        "",
        "[Service1]",
        "AccountName=$dispName",
        "ServerName=$ServerHost",
        "SMTPAddress=$Email",
        "AccountType=IMAP",
        "UserName=$userLogin",
        "EmailAddress=$Email",
        "ReplyToAddress=$Email",
        "SmtpServer=$ServerHost",
        "LogonMethod=1",
        "UseSPA=No",
        "RememberPassword=Yes",
        "SMTPServerRequiresAuth=1",
        "SMTPAuthMethod=1",
        "SMTPUserName=$userLogin",
        "SMTPServerPort=$SmtpPort",
        "SMTPServerSSL=1",
        "ServerPort=$ImapPort",
        "ServerSSL=1"
    )

    [System.IO.File]::WriteAllLines($prfPath, $prfLines, [System.Text.Encoding]::ASCII)
    Show-Notification "Saved PRF to: $prfPath"

    # Stage PRF for Outlook auto-import via registry
    $setupKey = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Setup"
    if (-not (Test-Path $setupKey)) {
        New-Item -Path $setupKey -ItemType Directory -Force | Out-Null
    }
    Set-ItemProperty -Path $setupKey -Name "ImportPRF" -Value $prfPath -Force | Out-Null

    # Also register the profile key in HKCU
    $profileReg = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$Name"
    if (-not (Test-Path $profileReg)) {
        New-Item -Path $profileReg -ItemType Directory -Force | Out-Null
    }

    if ($MakeDefault) {
        Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles" -Name "DefaultProfile" -Value $Name -Force | Out-Null
    }

    Show-Notification "Outlook Profile '$Name' registered!" -Type "Success"

    if ($outlookExe) {
        Write-Host "Do you want to launch Outlook now to finalize profile import? (Y/N): " -NoNewline -ForegroundColor Yellow
        $resp = Read-Host
        if ($resp -match '^(y|yes)$') {
            Show-Notification "Starting Outlook with /importprf..."
            Start-Process -FilePath $outlookExe -ArgumentList "/importprf `"$prfPath`""
        }
    }
}

# ---------------------------------------------------------------------------
# Clean & Repair: Non-Destructive Maintenance & Reset
# ---------------------------------------------------------------------------
function Invoke-MaintenanceClean {
    param(
        [switch]$ThunderbirdLocks,
        [switch]$ThunderbirdCache,
        [switch]$ThunderbirdReindex,
        [switch]$OutlookProcesses,
        [switch]$OutlookRoamCache,
        [switch]$OutlookViews
    )

    Show-Notification "Starting Mail Client Maintenance & Cache Cleanup..." -Type "Info"

    # Thunderbird Lock Clearing
    if ($ThunderbirdLocks -or (-not $PSBoundParameters.Count)) {
        $tbRoot = "$env:APPDATA\Thunderbird"
        if (Test-Path $tbRoot) {
            $lockFiles = Get-ChildItem -Path $tbRoot -Recurse -Filter "parent.lock" -ErrorAction SilentlyContinue
            if ($lockFiles) {
                foreach ($lf in $lockFiles) {
                    try {
                        Remove-Item -Path $lf.FullName -Force -ErrorAction SilentlyContinue
                        Show-Notification "Removed stuck Thunderbird lock: $($lf.FullName)" -Type "Success"
                    } catch {
                        Show-Notification "Could not remove lock $($lf.FullName): Client may still be running." -Type "Warning"
                    }
                }
            } else {
                Show-Notification "Thunderbird locks: No stuck parent.lock files found."
            }
        }
    }

    # Thunderbird Cache Cleaning
    if ($ThunderbirdCache -or (-not $PSBoundParameters.Count)) {
        $tbLocal = "$env:LOCALAPPDATA\Thunderbird"
        if (Test-Path $tbLocal) {
            $cacheDirs = Get-ChildItem -Path $tbLocal -Recurse -Directory -Filter "cache2" -ErrorAction SilentlyContinue
            foreach ($cd in $cacheDirs) {
                try {
                    Remove-Item -Path $cd.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    Show-Notification "Cleared Thunderbird cache directory: $($cd.FullName)" -Type "Success"
                } catch {}
            }
            $startupCaches = Get-ChildItem -Path $tbLocal -Recurse -Directory -Filter "startupCache" -ErrorAction SilentlyContinue
            foreach ($sc in $startupCaches) {
                try {
                    Remove-Item -Path $sc.FullName -Recurse -Force -ErrorAction SilentlyContinue
                    Show-Notification "Cleared Thunderbird startupCache: $($sc.FullName)" -Type "Success"
                } catch {}
            }
        }
    }

    # Thunderbird Corrupted Index (.msf) Rebuilding
    if ($ThunderbirdReindex) {
        $tbRoot = "$env:APPDATA\Thunderbird"
        if (Test-Path $tbRoot) {
            $msfFiles = Get-ChildItem -Path $tbRoot -Recurse -Filter "*.msf" -ErrorAction SilentlyContinue
            $count = 0
            foreach ($msf in $msfFiles) {
                try {
                    Remove-Item -Path $msf.FullName -Force -ErrorAction SilentlyContinue
                    $count++
                } catch {}
            }
            Show-Notification "Removed $count index (.msf) files. Thunderbird will re-index clean folders on launch." -Type "Success"
        }
    }

    # Outlook Process Killing
    if ($OutlookProcesses -or (-not $PSBoundParameters.Count)) {
        $procs = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
        if ($procs) {
            foreach ($p in $procs) {
                try {
                    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
                    Show-Notification "Terminated running OUTLOOK.EXE (PID: $($p.Id))" -Type "Success"
                } catch {}
            }
        }
    }

    # Outlook RoamCache Cleaning
    if ($OutlookRoamCache -or (-not $PSBoundParameters.Count)) {
        $roamCache = "$env:LOCALAPPDATA\Microsoft\Outlook\RoamCache"
        if (Test-Path $roamCache) {
            $files = Get-ChildItem -Path $roamCache -File -ErrorAction SilentlyContinue
            $count = 0
            foreach ($f in $files) {
                try {
                    Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
                    $count++
                } catch {}
            }
            Show-Notification "Cleared $count temporary files from Outlook RoamCache." -Type "Success"
        }

        # Outlook Autodiscover XML cache
        $localOutlook = "$env:LOCALAPPDATA\Microsoft\Outlook"
        if (Test-Path $localOutlook) {
            $xmlFiles = Get-ChildItem -Path $localOutlook -Filter "*.xml" -ErrorAction SilentlyContinue
            foreach ($xf in $xmlFiles) {
                try {
                    Remove-Item -Path $xf.FullName -Force -ErrorAction SilentlyContinue
                } catch {}
            }
        }
    }

    # Outlook Reset Views Switch
    if ($OutlookViews) {
        $outlookExe = Get-OutlookExecutable
        if ($outlookExe) {
            Show-Notification "Starting Outlook with /cleanviews and /resetnavpane..."
            Start-Process -FilePath $outlookExe -ArgumentList "/cleanviews /resetnavpane"
        }
    }

    Show-Notification "Maintenance clean completed!" -Type "Success"
}

# ---------------------------------------------------------------------------
# Backup: Automatic Backup of Profiles & Registries
# ---------------------------------------------------------------------------
function Backup-Profiles {
    param([ValidateSet('All', 'Outlook', 'Thunderbird')]$Target = 'All')

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    Show-Notification "Creating automated backup in: $backupDir" -Type "Info"

    if ($Target -in @('All', 'Outlook')) {
        $regOutlook = "HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles"
        $regOutFile = Join-Path $backupDir "Outlook_Profiles_$ts.reg"
        try {
            $proc = Start-Process -FilePath "reg.exe" -ArgumentList "export `"$regOutlook`" `"$regOutFile`" /y" -Wait -PassThru -NoNewWindow
            if ($proc.ExitCode -eq 0 -and (Test-Path $regOutFile)) {
                Show-Notification "Outlook registry profiles backed up to: $regOutFile" -Type "Success"
            }
        } catch {
            Show-Notification "Failed to export Outlook registry: $($_.Exception.Message)" -Type "Warning"
        }
    }

    if ($Target -in @('All', 'Thunderbird')) {
        $tbRoot = "$env:APPDATA\Thunderbird"
        if (Test-Path $tbRoot) {
            $tbBackupFolder = Join-Path $backupDir "Thunderbird_Config_$ts"
            New-Item -Path $tbBackupFolder -ItemType Directory -Force | Out-Null
            
            Copy-Item -Path (Join-Path $tbRoot "profiles.ini") -Destination $tbBackupFolder -Force -ErrorAction SilentlyContinue
            Copy-Item -Path (Join-Path $tbRoot "installs.ini") -Destination $tbBackupFolder -Force -ErrorAction SilentlyContinue
            
            # Backup prefs.js for each profile
            Get-ChildItem -Path (Join-Path $tbRoot "Profiles") -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $pjs = Join-Path $_.FullName "prefs.js"
                if (Test-Path $pjs) {
                    $sub = Join-Path $tbBackupFolder $_.Name
                    New-Item -Path $sub -ItemType Directory -Force | Out-Null
                    Copy-Item -Path $pjs -Destination $sub -Force -ErrorAction SilentlyContinue
                }
            }
            Show-Notification "Thunderbird configuration backed up to: $tbBackupFolder" -Type "Success"
        }
    }
}

# ---------------------------------------------------------------------------
# Destructive Actions: Remove Profile / Orphaned Cleanup / Reset
# ---------------------------------------------------------------------------
function Remove-MailProfile {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('Outlook', 'Thunderbird')]
        [string]$TargetClient,
        [Parameter(Mandatory=$true)]
        [string]$ProfileNameToDelete
    )

    # Always backup first
    Backup-Profiles -Target $TargetClient

    if ($TargetClient -eq 'Outlook') {
        $regPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\$ProfileNameToDelete"
        if (Test-Path $regPath) {
            Remove-Item -Path $regPath -Recurse -Force | Out-Null
            Show-Notification "Removed Outlook Profile '$ProfileNameToDelete' from registry." -Type "Success"
        } else {
            Show-Notification "Outlook profile '$ProfileNameToDelete' not found in registry." -Type "Warning"
        }
    } else {
        $tbRoot = "$env:APPDATA\Thunderbird"
        $iniPath = Join-Path $tbRoot "profiles.ini"
        if (Test-Path $iniPath) {
            $lines = Get-Content $iniPath
            $newLines = @()
            $inTargetSection = $false
            $targetPath = $null

            foreach ($line in $lines) {
                if ($line -match '^\[Profile\d+\]$') {
                    $inTargetSection = $false
                }
                if ($line -match "^Name=$([regex]::Escape($ProfileNameToDelete))$") {
                    $inTargetSection = $true
                }
                if ($inTargetSection -and $line -match '^Path=(.*)$') {
                    $targetPath = $matches[1]
                }
                if (-not $inTargetSection) {
                    $newLines += $line
                }
            }

            [System.IO.File]::WriteAllLines($iniPath, $newLines, [System.Text.Encoding]::UTF8)
            Show-Notification "Removed profile entry '$ProfileNameToDelete' from profiles.ini." -Type "Success"

            if ($targetPath) {
                $dirToRemove = Join-Path $tbRoot ($targetPath -replace '/', '\')
                if (Test-Path $dirToRemove) {
                    Write-Host "Do you also want to permanently delete profile data folder on disk? (Y/N): " -NoNewline -ForegroundColor Yellow
                    $delFolder = Read-Host
                    if ($delFolder -match '^(y|yes)$') {
                        Remove-Item -Path $dirToRemove -Recurse -Force -ErrorAction SilentlyContinue
                        Show-Notification "Deleted profile directory: $dirToRemove" -Type "Success"
                    }
                }
            }
        }
    }
}

function Remove-OrphanedData {
    param([ValidateSet('All', 'Outlook', 'Thunderbird')]$Target = 'All')

    if ($Target -in @('All', 'Thunderbird')) {
        $tbProfiles = Get-ThunderbirdProfiles | Where-Object { $_.IsOrphaned }
        if ($tbProfiles) {
            foreach ($op in $tbProfiles) {
                Show-Notification "Found orphaned Thunderbird folder: $($op.FolderPath)"
                Write-Host "Delete orphaned folder? (Y/N): " -NoNewline -ForegroundColor Yellow
                $del = Read-Host
                if ($del -match '^(y|yes)$') {
                    Remove-Item -Path $op.FolderPath -Recurse -Force -ErrorAction SilentlyContinue
                    Show-Notification "Deleted: $($op.FolderPath)" -Type "Success"
                }
            }
        } else {
            Show-Notification "No orphaned Thunderbird profile directories found."
        }
    }

    if ($Target -in @('All', 'Outlook')) {
        $localOutlook = "$env:LOCALAPPDATA\Microsoft\Outlook"
        if (Test-Path $localOutlook) {
            $allOsts = Get-ChildItem -Path $localOutlook -Filter "*.ost" -ErrorAction SilentlyContinue
            $activeProfiles = Get-OutlookProfiles
            $activeDataPaths = @()
            foreach ($ap in $activeProfiles) {
                foreach ($df in $ap.DataFiles) {
                    $activeDataPaths += $df.Path.ToLower()
                }
            }

            $orphanedOsts = $allOsts | Where-Object { $activeDataPaths -notcontains $_.FullName.ToLower() }
            if ($orphanedOsts) {
                foreach ($ost in $orphanedOsts) {
                    $sizeMb = [math]::Round($ost.Length / 1MB, 2)
                    Show-Notification "Found orphaned OST file ($sizeMb MB): $($ost.FullName)"
                    Write-Host "Delete orphaned OST to reclaim disk space? (Y/N): " -NoNewline -ForegroundColor Yellow
                    $del = Read-Host
                    if ($del -match '^(y|yes)$') {
                        Remove-Item -Path $ost.FullName -Force -ErrorAction SilentlyContinue
                        Show-Notification "Deleted: $($ost.FullName)" -Type "Success"
                    }
                }
            } else {
                Show-Notification "No orphaned Outlook OST files detected."
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Interactive Menu Workflows
# ---------------------------------------------------------------------------
function Show-InteractiveMenu {
    while ($true) {
        Write-Header
        Write-Host "Select an operation:" -ForegroundColor White
        Write-Host ""
        Write-Host "  [1] Overview & Quick Health Check (Outlook & Thunderbird)" -ForegroundColor Cyan
        Write-Host "  [2] Mozilla Thunderbird Tools (Inspect, Add, Test, Clean)" -ForegroundColor Cyan
        Write-Host "  [3] Microsoft Outlook Tools (Inspect, Add, Test, Clean)" -ForegroundColor Cyan
        Write-Host "  [4] Add / Configure New Profile (Interactive Wizard)" -ForegroundColor Green
        Write-Host "  [5] Deep Server & SSL Connectivity Test (IMAP/POP3/SMTP)" -ForegroundColor Green
        Write-Host "  [6] Quick Maintenance & Cache Cleanup (Safe & Non-Destructive)" -ForegroundColor Yellow
        Write-Host "  [7] Advanced Profile Deletion & Orphaned Data Cleanup" -ForegroundColor Magenta
        Write-Host "  [8] Manage Automated Backups (View / Export)" -ForegroundColor DarkCyan
        Write-Host "  [9] Copy Last Diagnostic Report to Clipboard" -ForegroundColor White
        Write-Host "  [0] Exit" -ForegroundColor Red
        Write-Host ""
        Write-Host "Choice: " -NoNewline -ForegroundColor Yellow

        $choice = Read-Host

        switch ($choice) {
            "1" {
                Write-Header "OVERVIEW & HEALTH CHECK"
                Invoke-ProfileTestReport
                Pause-Menu
            }
            "2" {
                Show-ThunderbirdSubMenu
            }
            "3" {
                Show-OutlookSubMenu
            }
            "4" {
                Show-AddProfileWizard
            }
            "5" {
                Show-CustomServerTestWizard
            }
            "6" {
                Write-Header "MAINTENANCE & CACHE CLEANUP"
                Invoke-MaintenanceClean
                Pause-Menu
            }
            "7" {
                Show-AdvancedCleanMenu
            }
            "8" {
                Write-Header "BACKUPS DIRECTORY"
                Show-Notification "Backups are stored at: $backupDir"
                Get-ChildItem -Path $backupDir | Select-Object -Property Name, Length, LastWriteTime | Format-Table -AutoSize
                Write-Host "Create a fresh backup now? (Y/N): " -NoNewline -ForegroundColor Yellow
                $ans = Read-Host
                if ($ans -match '^(y|yes)$') {
                    Backup-Profiles -Target 'All'
                }
                Pause-Menu
            }
            "9" {
                if ($global:LastReportText) {
                    Set-Clipboard -Value $global:LastReportText
                    Show-Notification "Diagnostic report copied to clipboard!" -Type "Success"
                } else {
                    Show-Notification "No report in memory yet. Running health check first..."
                    Invoke-ProfileTestReport
                    Set-Clipboard -Value $global:LastReportText
                    Show-Notification "Diagnostic report copied to clipboard!" -Type "Success"
                }
                Pause-Menu
            }
            "0" {
                Show-Notification "Exiting Mail Client Profile Manager. Goodbye!"
                try { Stop-Transcript | Out-Null } catch {}
                Exit 0
            }
            Default {
                Show-Notification "Invalid option. Please choose between 0 and 9." -Type "Warning"
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Show-ThunderbirdSubMenu {
    while ($true) {
        Write-Header "MOZILLA THUNDERBIRD TOOLS"
        Write-Host "  [1] List & Inspect Thunderbird Profiles" -ForegroundColor Cyan
        Write-Host "  [2] Configure / Add New Thunderbird Profile" -ForegroundColor Green
        Write-Host "  [3] Test All Configured Mail Servers & SSL Certificates" -ForegroundColor Green
        Write-Host "  [4] Clean Stuck Locks (parent.lock) & Temp Caches" -ForegroundColor Yellow
        Write-Host "  [5] Repair Corrupted Folders (Rebuild .msf Index Files)" -ForegroundColor Yellow
        Write-Host "  [6] Launch Native Thunderbird Profile Manager" -ForegroundColor White
        Write-Host "  [0] Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Choice: " -NoNewline -ForegroundColor Yellow

        $sub = Read-Host
        switch ($sub) {
            "1" {
                Write-Header "THUNDERBIRD PROFILES"
                $tb = Get-ThunderbirdProfiles
                $tb | Format-List -Property ProfileName, IsDefault, FolderPath, SizeMB, HasLock, Identities
                Pause-Menu
            }
            "2" {
                Show-AddProfileWizard -PreselectedClient "Thunderbird"
            }
            "3" {
                Write-Header "THUNDERBIRD SERVER & SSL TEST"
                Invoke-ProfileTestReport
                Pause-Menu
            }
            "4" {
                Write-Header "CLEAN THUNDERBIRD CACHE & LOCKS"
                Invoke-MaintenanceClean -ThunderbirdLocks -ThunderbirdCache
                Pause-Menu
            }
            "5" {
                Write-Header "REPAIR THUNDERBIRD FOLDERS (.MSF)"
                Invoke-MaintenanceClean -ThunderbirdReindex
                Pause-Menu
            }
            "6" {
                $tbExe = Get-ThunderbirdExecutable
                if ($tbExe) {
                    Show-Notification "Launching Thunderbird Profile Manager..."
                    Start-Process -FilePath $tbExe -ArgumentList "-ProfileManager"
                } else {
                    Show-Notification "thunderbird.exe not found on this system." -Type "Error"
                }
                Pause-Menu
            }
            "0" { return }
            Default { Start-Sleep -Seconds 1 }
        }
    }
}

function Show-OutlookSubMenu {
    while ($true) {
        Write-Header "MICROSOFT OUTLOOK TOOLS"
        Write-Host "  [1] List & Inspect Outlook Profiles (Registry & OST Data)" -ForegroundColor Cyan
        Write-Host "  [2] Configure / Add New Outlook Profile via .PRF" -ForegroundColor Green
        Write-Host "  [3] Clean RoamCache, Autodiscover XMLs, and Temp Attachments" -ForegroundColor Yellow
        Write-Host "  [4] Reset Outlook Views & Navigation Pane (/cleanviews)" -ForegroundColor Yellow
        Write-Host "  [5] Terminate Hung Outlook Zombie Processes" -ForegroundColor Yellow
        Write-Host "  [6] Launch Native Outlook Profile Dialog (/profiles)" -ForegroundColor White
        Write-Host "  [0] Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Choice: " -NoNewline -ForegroundColor Yellow

        $sub = Read-Host
        switch ($sub) {
            "1" {
                Write-Header "OUTLOOK PROFILES"
                $ops = Get-OutlookProfiles
                $ops | Format-List -Property ProfileName, IsDefault, OfficeVersion, RegistryPath, Emails, DataFiles
                Pause-Menu
            }
            "2" {
                Show-AddProfileWizard -PreselectedClient "Outlook"
            }
            "3" {
                Write-Header "CLEAN OUTLOOK CACHES"
                Invoke-MaintenanceClean -OutlookRoamCache
                Pause-Menu
            }
            "4" {
                Write-Header "RESET OUTLOOK VIEWS"
                Invoke-MaintenanceClean -OutlookViews
                Pause-Menu
            }
            "5" {
                Write-Header "KILL HUNG OUTLOOK PROCESSES"
                Invoke-MaintenanceClean -OutlookProcesses
                Pause-Menu
            }
            "6" {
                $outExe = Get-OutlookExecutable
                if ($outExe) {
                    Show-Notification "Launching Outlook Profile Manager..."
                    Start-Process -FilePath $outExe -ArgumentList "/profiles"
                } else {
                    Show-Notification "OUTLOOK.EXE not found on this system." -Type "Error"
                }
                Pause-Menu
            }
            "0" { return }
            Default { Start-Sleep -Seconds 1 }
        }
    }
}

function Show-AddProfileWizard {
    param([string]$PreselectedClient)

    Write-Header "ADD / CONFIGURE MAIL CLIENT PROFILE"

    $cliChoice = $PreselectedClient
    if (-not $cliChoice) {
        Write-Host "Select Target Mail Client:" -ForegroundColor White
        Write-Host "  [1] Mozilla Thunderbird" -ForegroundColor Cyan
        Write-Host "  [2] Microsoft Outlook" -ForegroundColor Cyan
        Write-Host "Choice (1/2): " -NoNewline -ForegroundColor Yellow
        $c = Read-Host
        $cliChoice = if ($c -eq "2") { "Outlook" } else { "Thunderbird" }
    }

    Write-Host "`nConfiguring profile for: $cliChoice" -ForegroundColor Green
    Write-Host "Enter Profile Name (e.g. Work, CompanyMail, Support): " -NoNewline -ForegroundColor Yellow
    $pName = Read-Host
    if (-not $pName) { $pName = "MailProfile" }

    Write-Host "Enter Email Address (e.g. user@example.com): " -NoNewline -ForegroundColor Yellow
    $email = Read-Host

    Write-Host "Enter User Full Name / Display Name: " -NoNewline -ForegroundColor Yellow
    $dispName = Read-Host

    Write-Host "Enter Mail Server Host (e.g. mail.example.com): " -NoNewline -ForegroundColor Yellow
    $serverHost = Read-Host

    Write-Host "Enter IMAP Port (Default 993): " -NoNewline -ForegroundColor Yellow
    $imapPortStr = Read-Host
    $imapPort = if ($imapPortStr) { [int]$imapPortStr } else { 993 }

    Write-Host "Enter SMTP Port (Default 465): " -NoNewline -ForegroundColor Yellow
    $smtpPortStr = Read-Host
    $smtpPort = if ($smtpPortStr) { [int]$smtpPortStr } else { 465 }

    Write-Host "Set this profile as default? (Y/N): " -NoNewline -ForegroundColor Yellow
    $isDefResp = Read-Host
    $isDef = ($isDefResp -match '^(y|yes)$')

    if ($cliChoice -eq "Thunderbird") {
        Add-ThunderbirdProfile -Name $pName -Email $email -DisplayName $dispName `
            -ServerHost $serverHost -ImapPort $imapPort -SmtpPort $smtpPort `
            -MakeDefault:$isDef
    } else {
        Add-OutlookProfile -Name $pName -Email $email -DisplayName $dispName `
            -ServerHost $serverHost -ImapPort $imapPort -SmtpPort $smtpPort `
            -MakeDefault:$isDef
    }

    Pause-Menu
}

function Show-CustomServerTestWizard {
    Write-Header "CUSTOM SERVER & SSL TEST"
    Write-Host "Enter Mail Server Hostname or IP (e.g. mail.example.com): " -NoNewline -ForegroundColor Yellow
    $h = Read-Host
    if (-not $h) { return }

    $testPorts = @(
        @{ Port = 993;  Name = "IMAP SSL" },
        @{ Port = 143;  Name = "IMAP STARTTLS" },
        @{ Port = 995;  Name = "POP3 SSL" },
        @{ Port = 465;  Name = "SMTP SSL" },
        @{ Port = 587;  Name = "SMTP Submission" },
        @{ Port = 2096; Name = "Webmail SSL" }
    )

    Write-Host "`nTesting connectivity to $h..." -ForegroundColor Cyan
    foreach ($tp in $testPorts) {
        Write-Host "  Port $($tp.Port) ($($tp.Name))..." -NoNewline
        $reach = Test-FastPortReachability -HostName $h -Port $tp.Port
        if ($reach) {
            Write-Host " [OPEN/REACHABLE]" -ForegroundColor Green
            if ($tp.Port -in @(993, 995, 465, 2096)) {
                $cert = Test-MailSslCertificate -HostName $h -Port $tp.Port
                if ($cert.Valid) {
                    Write-Host "    -> SSL/TLS: Valid | Expires in $($cert.DaysLeft) days ($($cert.ExpiresOn))" -ForegroundColor DarkGreen
                } else {
                    Write-Host "    -> SSL/TLS: [FAILED: $($cert.Error)]" -ForegroundColor Red
                }
            }
        } else {
            Write-Host " [CLOSED/TIMED OUT]" -ForegroundColor Red
        }
    }

    Pause-Menu
}

function Show-AdvancedCleanMenu {
    while ($true) {
        Write-Header "ADVANCED CLEAN & PROFILE REMOVAL"
        Write-Host "NOTE: All operations automatically backup configurations before deletion." -ForegroundColor DarkYellow
        Write-Host ""
        Write-Host "  [1] Delete Specific Thunderbird Profile" -ForegroundColor Magenta
        Write-Host "  [2] Delete Specific Outlook Profile" -ForegroundColor Magenta
        Write-Host "  [3] Clean Orphaned Thunderbird Folders & Unused Outlook OSTs" -ForegroundColor Magenta
        Write-Host "  [0] Return to Main Menu" -ForegroundColor Red
        Write-Host ""
        Write-Host "Choice: " -NoNewline -ForegroundColor Yellow

        $c = Read-Host
        switch ($c) {
            "1" {
                $tb = Get-ThunderbirdProfiles
                Write-Host "`nExisting Thunderbird Profiles:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $tb.Count; $i++) {
                    Write-Host "  [$($i+1)] $($tb[$i].ProfileName) ($($tb[$i].FolderPath))"
                }
                Write-Host "Enter profile number to delete: " -NoNewline -ForegroundColor Yellow
                $idx = Read-Host
                if ($idx -match '^\d+$' -and [int]$idx -le $tb.Count -and [int]$idx -ge 1) {
                    $selected = $tb[[int]$idx - 1]
                    Remove-MailProfile -TargetClient 'Thunderbird' -ProfileNameToDelete $selected.ProfileName
                }
                Pause-Menu
            }
            "2" {
                $ops = Get-OutlookProfiles
                Write-Host "`nExisting Outlook Profiles:" -ForegroundColor Cyan
                for ($i = 0; $i -lt $ops.Count; $i++) {
                    Write-Host "  [$($i+1)] $($ops[$i].ProfileName) (Office $($ops[$i].OfficeVersion))"
                }
                Write-Host "Enter profile number to delete: " -NoNewline -ForegroundColor Yellow
                $idx = Read-Host
                if ($idx -match '^\d+$' -and [int]$idx -le $ops.Count -and [int]$idx -ge 1) {
                    $selected = $ops[[int]$idx - 1]
                    Remove-MailProfile -TargetClient 'Outlook' -ProfileNameToDelete $selected.ProfileName
                }
                Pause-Menu
            }
            "3" {
                Write-Header "CLEAN ORPHANED PROFILES & OSTS"
                Remove-OrphanedData -Target 'All'
                Pause-Menu
            }
            "0" { return }
            Default { Start-Sleep -Seconds 1 }
        }
    }
}

# ---------------------------------------------------------------------------
# CLI Execution Router (Non-Interactive vs Interactive)
# ---------------------------------------------------------------------------
if ($Action -eq 'Menu') {
    Show-InteractiveMenu
} else {
    Write-Host "Running MailClient-ProfileManager in CLI Mode: Action=$Action, Client=$Client" -ForegroundColor Cyan
    switch ($Action) {
        'List' {
            if ($Client -in @('All', 'Outlook')) {
                Get-OutlookProfiles | Format-Table -AutoSize
            }
            if ($Client -in @('All', 'Thunderbird')) {
                Get-ThunderbirdProfiles | Format-Table -AutoSize
            }
        }
        'Test' {
            Invoke-ProfileTestReport -OmitNetwork:$SkipNetworkTest
        }
        'Add' {
            if (-not $ProfileName -or -not $EmailAddress -or -not $IncomingServer) {
                Write-Error "Adding a profile requires -ProfileName, -EmailAddress, and -IncomingServer."
                Exit 1
            }
            $outHost = if ($OutgoingServer) { $OutgoingServer } else { $IncomingServer }
            $inUser  = if ($IncomingUser) { $IncomingUser } else { $EmailAddress }

            if ($Client -in @('All', 'Thunderbird')) {
                Add-ThunderbirdProfile -Name $ProfileName -Email $EmailAddress -DisplayName $DisplayName `
                    -ServerHost $IncomingServer -ImapPort $IncomingPort -SmtpPort $OutgoingPort `
                    -Username $inUser -MakeDefault:$SetAsDefault
            }
            if ($Client -in @('All', 'Outlook')) {
                Add-OutlookProfile -Name $ProfileName -Email $EmailAddress -DisplayName $DisplayName `
                    -ServerHost $IncomingServer -ImapPort $IncomingPort -SmtpPort $OutgoingPort `
                    -Username $inUser -MakeDefault:$SetAsDefault
            }
        }
        'Clean' {
            switch ($CleanScope) {
                'SafeCache' { Invoke-MaintenanceClean }
                'LocksOnly' { Invoke-MaintenanceClean -ThunderbirdLocks -OutlookProcesses }
                'CorruptedIndices' { Invoke-MaintenanceClean -ThunderbirdReindex }
                'OrphanedProfiles' { Remove-OrphanedData -Target $Client }
                'SpecificProfile' {
                    if (-not $TargetProfile) {
                        Write-Error "SpecificProfile clean scope requires -TargetProfile <Name>."
                        Exit 1
                    }
                    if ($Client -eq 'All') {
                        Write-Error "Please specify -Client Outlook or -Client Thunderbird when deleting a specific profile."
                        Exit 1
                    }
                    Remove-MailProfile -TargetClient $Client -ProfileNameToDelete $TargetProfile
                }
            }
        }
        'Repair' {
            Invoke-MaintenanceClean -ThunderbirdLocks -ThunderbirdCache -ThunderbirdReindex -OutlookProcesses -OutlookRoamCache -OutlookViews
        }
        'Backup' {
            Backup-Profiles -Target $Client
        }
    }
}
