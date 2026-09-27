# Project Handoff: IT Support Tools

**Last Updated:** September 2026
**Current Focus:** `Universal-MailRepair.ps1` (v1.1.7)

## Overview
This repository (`samir-koulali/it-support-tools`) houses diagnostic and repair scripts designed for IT technicians. The primary asset currently is a robust Windows PowerShell script built to seamlessly diagnose, repair, and rollback DNS and IPv6 routing issues that commonly break mail clients (especially against strict ISPs).

## 🏆 Accomplishments & Current State

### 1. The Windows Universal Mail Repair Tool (`v1.1.7`)
*   **Location:** `windows/Universal-MailRepair.ps1`
*   **Documentation:** `windows/Universal-MailRepair.md`
*   **Interactive Menu Loop:** The script acts as a persistent CLI application. After running any diagnostic or repair task, users are gracefully returned to the main menu (0 to exit, 1 to return).
*   **Deep Diagnostics:**
    *   **DNS Inspector:** Extracts A, AAAA, CNAME, MX, NS, SPF/TXT, DMARC, and AutoDiscover configurations.
    *   **Network Status:** Filters out virtual/VPN adapters and exposes actual physical adapter states, IPv4 DNS, and IPv6 bindings.
    *   **Port Reachability:** Custom, rapid C# `TcpClient` logic tests IMAP (993), POP3 (995), SMTP (465, 587), Webmail (2096), cPanel (2083), and HTTPS (443).
    *   **SSL/TLS Validation:** Inspects certificate expiration and Subject Alternative Names (SAN). Smart fallback logic prioritizes IMAPS (Port 993) over HTTPS (Port 443) for domains starting with `mail.*` to avoid false-negative mismatches.
*   **Intelligent Repair & Rollback:**
    *   Flushes DNS cache, assigns safe global DNS (Cloudflare/Google/Quad9/OpenDNS), disables the IPv6 binding on physical adapters, and safely updates the global Windows IPv6 prefix policy to prefer IPv4 (`netsh`).
    *   Records previous states (DHCP/Static, IPs, IPv6 states, and importantly, **Adapter MAC Addresses**) to a JSON file (`dns_backup_latest.json`) for bulletproof rollbacks, even if adapters are renamed or disconnected.

## ⚠️ Critical Edge Cases Handled (Do Not Revert)

1.  **The `Invoke-Expression` (iex) Parsing Bug:** 
    *   Running complex scripts directly from GitHub via `irm | iex` caused severe parsing failures when the script contained inline semicolons, nested subexpressions `$()`, and single quotes. 
    *   *Solution:* We refactored strings to use double quotes and strictly standardized Windows `CRLF` line endings. More importantly, we updated the official documentation to force users to download the script to `$env:TEMP` and execute it natively to completely bypass `iex` memory limits.
2.  **`Stop-Transcript` Terminating Errors:** 
    *   Because the tool loops infinitely, `Stop-Transcript` was throwing terminating exceptions when it tried to stop an already-stopped session.
    *   *Solution:* `Stop-Transcript` is wrapped in `try/catch` and strictly bound to the final `Exit 0` blocks.
3.  **TCP Socket DNS Hangs:** 
    *   The .NET `TcpClient` relies on Windows DNS. If the system DNS is broken, the port checks hung indefinitely.
    *   *Solution:* The diagnostic loop caches the IP resolved by the script's `Resolve-DnsName` check and feeds the raw IP into the Socket connection.

## 🚀 Next Steps & Future Ideas

*   **Expand Catalog:** Begin populating the `macos/` and `linux/` directories with equivalent bash/python tools.
*   **Mail Client Profiler:** Add functionality to inspect local Outlook profiles (via registry) or Thunderbird `prefs.js` to flag corrupt local profiles.
*   **Advanced DNS Propagation:** Integrate an API check (like DNSChecker or Google DoH) to verify if the local DNS matches the global DNS propagation state.
