# IT Support Tools

A collection of diagnostic, configuration, and repair scripts for IT support, systems administrators, helpdesks, and technicians. Scripts are organized by operating system and can be executed directly from GitHub via PowerShell without manual cloning or downloading.

---

## 🚀 Quick Start (Run Directly from Cloud)

You can run any tool instantly in PowerShell by copying and pasting the corresponding one-liner:

### 1. Mail Client Profile Manager (Outlook & Thunderbird)
Diagnose, configure, test live ports/SSL certificates, clean stuck locks, and fix corrupt profiles.

```powershell
irm "https://raw.githubusercontent.com/samir-koulali/it-support-tools/master/windows/MailClient-ProfileManager.ps1" -OutFile "$env:TEMP\MailProfileManager.ps1"; & "$env:TEMP\MailProfileManager.ps1"
```

* **Instant Diagnostic Test:**
  ```powershell
  irm "https://raw.githubusercontent.com/samir-koulali/it-support-tools/master/windows/MailClient-ProfileManager.ps1" -OutFile "$env:TEMP\MailProfileManager.ps1"; & "$env:TEMP\MailProfileManager.ps1" -Action Test
  ```
* **Silent Cache & Lock Cleanup:**
  ```powershell
  irm "https://raw.githubusercontent.com/samir-koulali/it-support-tools/master/windows/MailClient-ProfileManager.ps1" -OutFile "$env:TEMP\MailProfileManager.ps1"; & "$env:TEMP\MailProfileManager.ps1" -Action Clean -CleanScope SafeCache
  ```

---

### 2. Universal Mail Connection & Network Repair
Fixes DNS and broken IPv6 routing issues that block mail servers (IMAP, POP3, SMTP, Webmail, cPanel).

```powershell
irm "https://raw.githubusercontent.com/samir-koulali/it-support-tools/master/windows/Universal-MailRepair.ps1" -OutFile "$env:TEMP\MailRepair.ps1"; & "$env:TEMP\MailRepair.ps1"
```

---

## 📂 Script Catalog

### Windows (`/windows`)

| Tool | Description | Documentation |
| :--- | :--- | :--- |
| **Mail Client Profile Manager** | Add, test, and clean profiles for Microsoft Outlook & Mozilla Thunderbird. Live IMAP/SMTP port & SSL validation, lock clearing, and index rebuilds. | [View Guide](windows/MailClient-ProfileManager.md) |
| **Universal Mail Connection Repair** | Fixes DNS/IPv6 routing blocking mail clients against strict ISPs. Includes port diagnostics, SSL checks, and safe rollback. | [View Guide](windows/Universal-MailRepair.md) |

### macOS (`/macos`)
*(Coming soon - tools and scripts for macOS environments)*

### Linux (`/linux`)
*(Coming soon - tools and scripts for Linux environments)*

---

**License:** Open source for non-commercial use.
