# Active Directory Lab Automation Toolkit

## Overview
The `AllInOne-AD-Setup.ps1` script automates the buildout of an Active Directory lab by prompting for environment details, creating the entire departmental folder/share structure, provisioning OUs and security groups, importing users, and applying NTFS/share permissions in an idempotent, rerunnable way.【F:AllInOne-AD-Setup.ps1†L1-L8】【F:AllInOne-AD-Setup.ps1†L157-L309】

## Key capabilities
- **GUI-driven configuration.** Windows Forms dialogs gather the share host, base path, department list, passwords, optional CSV path, UPN suffix, and alternate admin details, falling back to parameters when provided for scripted runs.【F:AllInOne-AD-Setup.ps1†L29-L160】【F:AllInOne-AD-Setup.ps1†L233-L309】
- **Automatic share and folder provisioning.** The script ensures `Home`, `Profiles`, and per-department directories exist, sharing each one with `Everyone` at the share level so reruns only report status changes.【F:AllInOne-AD-Setup.ps1†L715-L730】
- **OU and group creation.** Every department receives a matching OU and global security group using resilient helper functions that log when items already exist.【F:AllInOne-AD-Setup.ps1†L732-L757】
- **Standardized NTFS permissions.** Department folders inherit SYSTEM/Administrators full control, IT full control, department-group full control, and Executives/Management read-only (skipping IT) while removing the default `BUILTIN\Users` entry.【F:AllInOne-AD-Setup.ps1†L802-L857】
- **CSV-driven user provisioning.** The importer handles `FullName` or `First/Last` columns, generates unique sAMAccountNames, randomizes department placement, forces password reset at next logon, and builds per-user home folders with ACLs.【F:AllInOne-AD-Setup.ps1†L870-L1010】
- **Alternate admin account creation.** A helper provisions or reuses a lab admin account, enforces group memberships (IT, Domain Admins, Administrators), and records the outcome for verification.【F:AllInOne-AD-Setup.ps1†L550-L683】
- **Run summary.** Post-run tables list shares, OUs, groups, user import counts, and alternate admin details for quick validation.【F:AllInOne-AD-Setup.ps1†L1026-L1082】

## Prerequisites
- Windows Server 2016 or later (PowerShell 5.1).
- Active Directory Domain Services and file services roles installed.
- PowerShell modules: **ActiveDirectory** and **SmbShare** (loaded automatically).【F:AllInOne-AD-Setup.ps1†L25-L27】
- Run the script from an elevated PowerShell ISE or console session.

## Running the script
1. Sign in to a domain controller (or management VM) with administrative privileges.
2. Copy `AllInOne-AD-Setup.ps1` to a convenient location.
3. Launch PowerShell ISE or Windows PowerShell **as Administrator**.
4. Execute the script directly:
   ```powershell
   Set-Location -Path "C:\Path\To\Script"
   .\AllInOne-AD-Setup.ps1
   ```
   - Optional switches: `-SkipFoldersAndShares` (only manage AD objects) and `-SkipAclForDepartments` (avoid NTFS resets).【F:AllInOne-AD-Setup.ps1†L21-L22】【F:AllInOne-AD-Setup.ps1†L715-L857】
5. Follow the GUI prompts (or pre-supply parameters) to complete configuration.

## GUI prompts and what they collect
| Prompt | Purpose |
| --- | --- |
| **Share Host** | Computer that will host SMB shares (defaults to the current computer).【F:AllInOne-AD-Setup.ps1†L157-L176】|
| **Base Share Path** | Root folder for `Home`, `Profiles`, and department directories (defaults to `C:\Shares`).【F:AllInOne-AD-Setup.ps1†L177-L204】|
| **Departments** | Comma-separated list (minimum five) used to build folders, OUs, and groups.【F:AllInOne-AD-Setup.ps1†L205-L232】【F:AllInOne-AD-Setup.ps1†L715-L757】|
| **Default Password** | Starting password for imported accounts; converted to a secure string for reuse.【F:AllInOne-AD-Setup.ps1†L233-L252】|
| **UPN Suffix** | Overrides the detected domain UPN suffix when needed (defaults to the AD DNS root).【F:AllInOne-AD-Setup.ps1†L253-L274】|
| **CSV Import** | Optional file picker for the user list; skip to build structure only.【F:AllInOne-AD-Setup.ps1†L275-L302】【F:AllInOne-AD-Setup.ps1†L870-L924】|
| **Alternate Admin (Display Name / Username / Password)** | Creates or updates a reusable admin account and secures it with the provided password.【F:AllInOne-AD-Setup.ps1†L303-L309】【F:AllInOne-AD-Setup.ps1†L550-L683】|

## CSV format guidelines
- Acceptable headers: `FullName` or `Name`, or combinations such as `FirstName`/`LastName`, `Given Name`/`Surname`, etc.【F:AllInOne-AD-Setup.ps1†L870-L914】
- Each valid row produces a unique sAMAccountName (first initial + last name with collision handling).【F:AllInOne-AD-Setup.ps1†L918-L947】
- Malformed entries are skipped and logged to `MalformedNames.txt` on the desktop for review.【F:AllInOne-AD-Setup.ps1†L1014-L1019】

## Folder, share, and permission model
1. **Folder creation** – Ensures the base path, `Home`, `Profiles`, and every department directory exist.【F:AllInOne-AD-Setup.ps1†L715-L730】
2. **Share creation** – Publishes each directory with `Everyone` Full Control at the share layer; reruns simply verify access.【F:AllInOne-AD-Setup.ps1†L446-L468】【F:AllInOne-AD-Setup.ps1†L715-L729】
3. **NTFS defaults** – Applies:
   - SYSTEM & Administrators: Full Control
   - IT group: Full Control on all folders
   - Department group: Full Control on its folder
   - Executives & Management: Read (skipped on IT folder)
   - Removes `BUILTIN\Users` from each directory.【F:AllInOne-AD-Setup.ps1†L802-L857】

## User provisioning workflow
- Places each new account in a random department OU and adds it to the department’s security group.【F:AllInOne-AD-Setup.ps1†L925-L970】
- Assigns `H:` drive mapping, creates `\\\\<Server>\\Home\\<sam>` directories, and applies user/IT/SYSTEM ACLs.【F:AllInOne-AD-Setup.ps1†L972-L1009】
- Forces `ChangePasswordAtLogon` for every imported or existing account the script touches.【F:AllInOne-AD-Setup.ps1†L934-L967】

## Alternate administrator account
- Prompts for display name, preferred username, and password to seed or reuse an elevated account.【F:AllInOne-AD-Setup.ps1†L303-L309】【F:AllInOne-AD-Setup.ps1†L550-L683】
- Places the account in the IT OU when available, enables it, and joins IT, Domain Admins, and Administrators groups with duplicate-safe logging.【F:AllInOne-AD-Setup.ps1†L612-L683】

## Verification and reporting
At completion the script prints a summary including:
- Share table filtered to the lab shares.【F:AllInOne-AD-Setup.ps1†L1026-L1031】
- OU listing for every supplied department.【F:AllInOne-AD-Setup.ps1†L1032-L1038】
- Department-to-group mapping with `<missing>` flags when unresolved.【F:AllInOne-AD-Setup.ps1†L1040-L1051】
- Alternate admin status and group membership notes.【F:AllInOne-AD-Setup.ps1†L1053-L1067】
- User import counts and malformed-name log location (if applicable).【F:AllInOne-AD-Setup.ps1†L1069-L1079】

## Troubleshooting tips
- Run the script from an elevated session to avoid SMB share permission denials.【F:AllInOne-AD-Setup.ps1†L446-L468】
- Ensure the IT, Executives, and Management departments are included in the prompt if their ACLs are required; missing groups trigger yellow status messages rather than failures.【F:AllInOne-AD-Setup.ps1†L824-L852】
- Review warnings about ACL updates or malformed CSV rows in the console output; the script continues processing but highlights any skipped actions.【F:AllInOne-AD-Setup.ps1†L350-L441】【F:AllInOne-AD-Setup.ps1†L1014-L1019】

## Need a scripted run?
All prompts correspond to parameters declared at the top of the script, letting you call it non-interactively (e.g., during lab resets) with custom values:
```powershell
.\u200bAllInOne-AD-Setup.ps1 -ServerName 'DC01' -BasePath 'D:\Shares' -Departments 'HR','IT','Accounting','Executives','Facilities' `
  -DefaultPassword 'P@ssw0rd!' -UPNSuffix 'lab.example.com' -CsvPath 'C:\Imports\users.csv' `
  -AltAdminDisplayName 'Lab Admin' -AltAdminSamAccountName 'labadmin' -AltAdminPassword 'Sup3rS3cret!'
```
Parameters can be mixed with switches like `-SkipFoldersAndShares` depending on what needs to be rebuilt.【F:AllInOne-AD-Setup.ps1†L10-L23】【F:AllInOne-AD-Setup.ps1†L715-L757】
