# Mail Client Profile Manager (Outlook & Thunderbird)

**Script:** `MailClient-ProfileManager.ps1`

A comprehensive diagnostic, configuration, testing, and cleanup utility designed for IT technicians, helpdesks, and system administrators. It automates managing mail client profiles for both **Microsoft Outlook** (Office 365, 2021, 2019, 2016, 2013) and **Mozilla Thunderbird**.

---

### Features & Capabilities

#### 1. 🔍 Deep Profile Discovery & Inspection
* **Microsoft Outlook:**
  * Detects installed Office versions (Click-to-Run, 64-bit, 32-bit).
  * Scans registry profiles under `HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles`.
  * Identifies the current `DefaultProfile`.
  * Extracts associated email accounts and server definitions.
  * Discovers all offline data files (`.ost` / `.pst`), computes exact sizes, and warns if any `.ost` file is exceeding the 45 GB threshold (approaching Outlook's 50 GB crash limit).
* **Mozilla Thunderbird:**
  * Parses `profiles.ini` and `installs.ini`.
  * Analyzes all profile folders on disk in `%APPDATA%\Thunderbird\Profiles\`.
  * Detects **orphaned profile folders** (folders wasting disk space without an entry in `profiles.ini`).
  * Deep-parses `prefs.js` to extract all configured email identities, incoming servers (IMAP/POP3), ports, usernames, and outgoing SMTP servers.
  * Flags lock states (`parent.lock`) caused by running instances or crashed processes.

#### 2. ⚡ Configure & Add Profiles
* **Thunderbird:**
  * Creates new profile folders with unique random 8-character salts.
  * Seeds a clean, pre-configured `user.js` with full account definitions (IMAP, SMTP, SSL/TLS, port 993/465, identities).
  * Automatically updates `profiles.ini` and sets the profile as active/default if requested.
* **Microsoft Outlook:**
  * Generates an official Microsoft Outlook Profile Configuration file (`.prf`).
  * Configures IMAP, SMTP, ports, SSL encryption flags, and credentials.
  * Stages the profile for instant import via the Outlook Setup registry or executes `OUTLOOK.EXE /importprf`.
  * Can also launch native profile creation wizards (`OUTLOOK.EXE /profiles`).

#### 3. 🧪 Live Port Reachability & SSL/TLS Verification
* Tests actual network connectivity from the client machine to every configured incoming and outgoing mail server.
* Supports **IMAP** (993, 143), **POP3** (995), and **SMTP** (465, 587).
* Validates SSL/TLS certificates: checks expiration date, days remaining, issuer, and flags expired or invalid certificates.
* Uses cached IP sockets to prevent indefinite hangs on broken local DNS.

#### 4. 🧹 Safe Maintenance, Cache Cleanup & Repair
* **Thunderbird Maintenance:**
  * **Unlock Profiles:** Clears stuck `parent.lock` files without rebooting.
  * **Cache Purge:** Cleans `cache2`, `startupCache`, and temporary data without affecting emails.
  * **Rebuild Folders (.msf):** Deletes corrupted folder index files (`*.msf`), prompting Thunderbird to re-index all folders safely on launch to fix missing/disappearing email bugs.
* **Outlook Maintenance:**
  * **Process Cleanup:** Safely terminates hung background `OUTLOOK.EXE` processes.
  * **RoamCache Reset:** Cleans `%LOCALAPPDATA%\Microsoft\Outlook\RoamCache` (fixes autocomplete and corrupted search/stream caches).
  * **Autodiscover XML Cache:** Cleans stale Autodiscover responses.
  * **View Reset:** Automates `/cleanviews` and `/resetnavpane` command-line switches.
* **Disk Space Reclamation:**
  * Scans for orphaned `.ost` files from deleted accounts and removes them.
  * Cleans orphaned Thunderbird profile directories.

#### 5. 🛡️ Bulletproof Backups
* **Automated Registry Backups:** Automatically exports Outlook registry profiles to `.reg` before any profile modification or deletion.
* **Thunderbird Backups:** Copies `profiles.ini`, `installs.ini`, and all `prefs.js` files to `$env:ProgramData\MailClientManager\backups\` before modifications.

---

### Quick Start (Interactive Menu)

Run the script in PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File "windows\MailClient-ProfileManager.ps1"
```

You will be greeted by the interactive console menu:

```text
==========================================================
   MAIL CLIENT PROFILE MANAGER (OUTLOOK & THUNDERBIRD)   
==========================================================
Select an operation:

  [1] Overview & Quick Health Check (Outlook & Thunderbird)
  [2] Mozilla Thunderbird Tools (Inspect, Add, Test, Clean)
  [3] Microsoft Outlook Tools (Inspect, Add, Test, Clean)
  [4] Add / Configure New Profile (Interactive Wizard)
  [5] Deep Server & SSL Connectivity Test (IMAP/POP3/SMTP)
  [6] Quick Maintenance & Cache Cleanup (Safe & Non-Destructive)
  [7] Advanced Profile Deletion & Orphaned Data Cleanup
  [8] Manage Automated Backups (View / Export)
  [9] Copy Last Diagnostic Report to Clipboard
  [0] Exit
```

---

### CLI & Automation Usage (Non-Interactive)

IT support teams and RMM tools (Datto, NinjaOne, Intune, ConnectWise) can run the script silently using parameters:

#### 1. Generate System Diagnostic Report
```powershell
.\MailClient-ProfileManager.ps1 -Action Test
```
*(Use `-SkipNetworkTest` to run local file/registry health checks only)*

#### 2. List Profiles
```powershell
# List profiles for both clients
.\MailClient-ProfileManager.ps1 -Action List -Client All

# List only Outlook profiles
.\MailClient-ProfileManager.ps1 -Action List -Client Outlook
```

#### 3. Run Non-Destructive Cache & Lock Cleanup
```powershell
# Clean safe caches and locks for both clients
.\MailClient-ProfileManager.ps1 -Action Clean -CleanScope SafeCache

# Remove stuck locks and terminate zombie mail processes
.\MailClient-ProfileManager.ps1 -Action Clean -CleanScope LocksOnly

# Repair corrupted Thunderbird folder indexes (.msf)
.\MailClient-ProfileManager.ps1 -Action Clean -CleanScope CorruptedIndices
```

#### 4. Add / Configure a Profile Silently
```powershell
.\MailClient-ProfileManager.ps1 -Action Add -Client Thunderbird `
    -ProfileName "Work" `
    -EmailAddress "user@company.com" `
    -DisplayName "John Doe" `
    -IncomingServer "mail.company.com" `
    -IncomingPort 993 `
    -OutgoingPort 465 `
    -SetAsDefault
```

#### 5. Create a Full Backup
```powershell
.\MailClient-ProfileManager.ps1 -Action Backup -Client All
```

---

### Data & Log Locations

* **Logs & Transcripts:** `C:\ProgramData\MailClientManager\logs\`
* **Registry & Config Backups:** `C:\ProgramData\MailClientManager\backups\`
* **Outlook Setup PRF Files:** `C:\ProgramData\MailClientManager\backups\Outlook_<ProfileName>.prf`
