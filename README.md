# IT Support Tools

A repository of diagnostic and repair scripts for IT support, systems administrators, and technicians.

---

## Windows Mail Connection Repair Tool
**Path:** `windows/Universal-MailRepair.ps1`

This tool was specifically designed to resolve connectivity issues where certain ISPs (such as those in Algeria) serve broken IPv6 addresses or randomly block IPv4 DNS resolution for mail servers. It restores reliable connectivity by bypassing the local router's DNS, enforcing reliable public DNS, and disabling broken IPv6 routing.

### Features
* **Interactive Menu:** Run the script and choose your operation mode on the fly.
* **Repair Mode:** Automatically disables IPv6 on active adapters and sets a custom DNS (Cloudflare, Google, Quad9, or OpenDNS).
* **Test Mode:** Runs diagnostics for DNS resolution, IMAP (993), and SMTP (465, 587).
* **Rollback Mode:** Safely undoes any network changes using an automated backup file.
* **View Status:** Quickly list all active physical network adapters and their current DNS settings.

### Quick Start (Run directly from the cloud)
You don't need to manually download the script to use it. You can run it directly from GitHub using PowerShell.

1. Click the Start Menu, type `PowerShell`.
2. Right-click **Windows PowerShell** and select **Run as Administrator**.
3. Copy and paste the following command, then press Enter:

```powershell
irm "https://raw.githubusercontent.com/samir-koulali/it-support-tools/master/windows/Universal-MailRepair.ps1" | iex
```

4. Follow the on-screen interactive menu!

#### Bypassing the menu (Advanced)
If you want to run the script in a specific mode without using the interactive menu (for example, triggering a rollback automatically), you can pass parameters by using a ScriptBlock:

```powershell
# Example: Triggering a rollback directly
& ([scriptblock]::Create((irm "https://raw.githubusercontent.com/samir-koulali/it-support-tools/master/windows/Universal-MailRepair.ps1"))) -Rollback
```

---
**License:** Open source for non-commercial use.
