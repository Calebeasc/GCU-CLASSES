# Active Directory Lab Automation Toolkit

## Overview
`AllInOne-AD-Setup.ps1` now ships as an “app style” PowerShell experience: a Windows Forms front-end collects lab-specific inputs, then the script provisions shares, OUs, groups, users, NTFS permissions, and an optional alternate administrator in an idempotent fashion.【F:AllInOne-AD-Setup.ps1†L1-L17】【F:AllInOne-AD-Setup.ps1†L872-L986】

## Key capabilities
- **Rich GUI workflow.** A configuration dialog gathers the share host, base path, department list, passwords, CSV file, UPN suffix, and alternate admin details; if parameters are supplied (or `-NoGui` is used) the same values are honoured for automated runs.【F:AllInOne-AD-Setup.ps1†L220-L405】【F:AllInOne-AD-Setup.ps1†L872-L895】
- **Live activity log window.** During interactive runs the script opens a resizable log window that mirrors console output and provides a Close button once provisioning completes.【F:AllInOne-AD-Setup.ps1†L95-L149】【F:AllInOne-AD-Setup.ps1†L981-L986】
- **Automatic share/folder creation.** Helper routines normalize paths, build the base, Home, Profiles, and per-department directories, and publish each share with `Everyone` Full Control while logging when items already exist.【F:AllInOne-AD-Setup.ps1†L155-L213】【F:AllInOne-AD-Setup.ps1†L412-L449】【F:AllInOne-AD-Setup.ps1†L932-L945】
- **OU and group provisioning.** Every department receives a matching OU and global security group; reruns simply report that the resources already exist.【F:AllInOne-AD-Setup.ps1†L451-L485】【F:AllInOne-AD-Setup.ps1†L946-L951】
- **Standard NTFS model.** Department folders gain SYSTEM/Administrators Full Control, department group Full Control, IT Full Control, and Executives/Management Read (except the IT folder) while removing `BUILTIN\Users`.【F:AllInOne-AD-Setup.ps1†L576-L614】【F:AllInOne-AD-Setup.ps1†L958-L967】
- **CSV-driven user import.** The importer accepts `FullName` or `First/Last` headers, generates unique sAMAccountNames, randomizes department placement, enforces password change at logon, and provisions home folders with ACLs.【F:AllInOne-AD-Setup.ps1†L664-L804】
- **Alternate admin creation.** A dedicated helper seeds or reuses an elevated account, ensuring membership in the IT, Domain Admins, and Administrators groups for lab administration.【F:AllInOne-AD-Setup.ps1†L620-L657】【F:AllInOne-AD-Setup.ps1†L970-L974】
- **Post-run verification.** A consolidated summary lists shares, OUs, groups, user import counts, malformed-name logs, and alternate admin status for quick validation.【F:AllInOne-AD-Setup.ps1†L810-L866】

## Prerequisites
- Windows Server 2016 or later (PowerShell 5.1).
- Active Directory Domain Services and file services roles installed.
- PowerShell modules: **ActiveDirectory** and **SmbShare** (loaded automatically).【F:AllInOne-AD-Setup.ps1†L40-L45】
- Run the script from an elevated PowerShell ISE or console session to create shares and write ACLs.【F:AllInOne-AD-Setup.ps1†L440-L449】

## Running the script
### Option 1 – Double-click launcher (no manual PowerShell window)
1. Copy `AllInOne-AD-Setup.ps1` **and** `Launch-AllInOne-AD-Setup.cmd` into the same folder (for example `C:\LabTools`).
2. Right-click `Launch-AllInOne-AD-Setup.cmd` and choose **Run as administrator** (or create a desktop shortcut to the CMD file and mark it to run elevated).
3. The launcher starts `powershell.exe` with the necessary execution policy flags and opens the GUI automatically. Complete the form and select **Run Provisioning**.【F:AllInOne-AD-Setup.ps1†L220-L405】【F:AllInOne-AD-Setup.ps1†L872-L979】

### Option 2 – Launch from PowerShell ISE/console
1. Sign in to a domain controller (or management workstation) with administrative privileges.
2. Launch PowerShell ISE or Windows PowerShell **as Administrator**.
3. Navigate to the script directory and run:
   ```powershell
   Set-Location -Path "C:\Path\To\Script"
   .\AllInOne-AD-Setup.ps1
   ```
4. Complete the GUI form and click **Run Provisioning**. Use parameters with `-NoGui` for unattended lab resets as needed.【F:AllInOne-AD-Setup.ps1†L220-L405】【F:AllInOne-AD-Setup.ps1†L872-L979】

> **Tip:** The base folder field accepts any valid path. If you enter a relative value (such as `Shares`) the script resolves it against the current working directory; leave the default `C:\Shares` or browse to an absolute location to avoid accidental placement under `C:\Windows\System32`.【F:AllInOne-AD-Setup.ps1†L155-L213】【F:AllInOne-AD-Setup.ps1†L903-L918】

## GUI prompt reference
| Prompt | Purpose |
| --- | --- |
| **Share host** | Computer that will host the SMB shares (defaults to the current computer).【F:AllInOne-AD-Setup.ps1†L249-L283】【F:AllInOne-AD-Setup.ps1†L379-L381】|
| **Base folder for shares** | Root folder for `Home`, `Profiles`, and department directories; includes a folder picker for convenience.【F:AllInOne-AD-Setup.ps1†L283-L295】|
| **Departments** | Comma or newline separated list (minimum of five) used for folders, shares, OUs, and groups.【F:AllInOne-AD-Setup.ps1†L296-L303】【F:AllInOne-AD-Setup.ps1†L372-L391】|
| **Default password** | Initial password assigned to imported accounts.【F:AllInOne-AD-Setup.ps1†L303-L307】【F:AllInOne-AD-Setup.ps1†L688-L743】|
| **UPN suffix** | Overrides the detected DNS root for account UPNs when required.【F:AllInOne-AD-Setup.ps1†L308-L310】【F:AllInOne-AD-Setup.ps1†L731-L734】|
| **User CSV path** | Optional file picker for the import list; leave blank to only build structure.【F:AllInOne-AD-Setup.ps1†L312-L320】【F:AllInOne-AD-Setup.ps1†L664-L704】|
| **Alternate admin (display, username, password)** | Seeds an additional administrative account that is granted IT and built-in admin group memberships.【F:AllInOne-AD-Setup.ps1†L321-L329】【F:AllInOne-AD-Setup.ps1†L620-L657】|
| **Skip options** | `Skip folder/share creation` and `Skip department ACL reset` switches mirror the script parameters for repeat runs.【F:AllInOne-AD-Setup.ps1†L331-L343】【F:AllInOne-AD-Setup.ps1†L872-L979】|

## CSV format guidelines
- Acceptable headers: `FullName` / `Name` or any of the supported `First` + `Last` combinations.【F:AllInOne-AD-Setup.ps1†L695-L704】
- Each row produces a unique sAMAccountName (first initial + last name with collision handling).【F:AllInOne-AD-Setup.ps1†L731-L754】
- Malformed entries are skipped and logged to `MalformedNames.txt` on the desktop for review.【F:AllInOne-AD-Setup.ps1†L795-L803】

## Folder, share, and permission model
1. **Folder creation** – Ensures the base path, `Home`, `Profiles`, and every department directory exist before provisioning NTFS permissions.【F:AllInOne-AD-Setup.ps1†L412-L424】【F:AllInOne-AD-Setup.ps1†L932-L944】
2. **Share creation** – Publishes each directory with `Everyone` Full Control and re-grants access on reruns.【F:AllInOne-AD-Setup.ps1†L427-L449】
3. **NTFS defaults** – Applies:
   - SYSTEM & Administrators: Full Control
   - Department security group: Full Control
   - IT security group: Full Control
   - Executives & Management: Read (skipped on the IT folder)
   - Removes `BUILTIN\Users` from each directory.【F:AllInOne-AD-Setup.ps1†L576-L614】【F:AllInOne-AD-Setup.ps1†L958-L967】

## User provisioning workflow
- Places each new account in a random department OU and adds the matching security group (plus Executives/Management when appropriate).【F:AllInOne-AD-Setup.ps1†L735-L770】
- Assigns `H:` drive mapping, creates `\\\\<Server>\\Home\\<sam>` directories, and sets ACLs for the user, IT, SYSTEM, and Administrators.【F:AllInOne-AD-Setup.ps1†L772-L791】
- Forces `ChangePasswordAtLogon` for created accounts while respecting reruns on existing users.【F:AllInOne-AD-Setup.ps1†L741-L755】

## Alternate administrator account
- Prompted details feed `Ensure-AlternateAdminAccount`, which creates or reuses the account inside the IT OU when present.【F:AllInOne-AD-Setup.ps1†L321-L329】【F:AllInOne-AD-Setup.ps1†L620-L657】
- The account is added to the IT group alongside the Domain Admins and Administrators built-in groups for elevated management tasks.【F:AllInOne-AD-Setup.ps1†L644-L655】

## Verification and reporting
After provisioning, review the summary for a quick health check:
- Share listing filtered to the lab shares.【F:AllInOne-AD-Setup.ps1†L827-L830】
- OU listing for every supplied department.【F:AllInOne-AD-Setup.ps1†L832-L838】
- Department-to-group mapping with `<missing>` markers when expected groups are absent.【F:AllInOne-AD-Setup.ps1†L840-L848】
- User import counts and malformed-name log location (if applicable).【F:AllInOne-AD-Setup.ps1†L851-L856】
- Alternate administrator account details when created.【F:AllInOne-AD-Setup.ps1†L859-L862】

## Troubleshooting tips
- Ensure the script is run elevated so SMB shares and ACL updates succeed.【F:AllInOne-AD-Setup.ps1†L440-L449】
- Include departments such as IT, Executives, and Management if their ACLs or groups are required; missing groups generate warnings but do not halt execution.【F:AllInOne-AD-Setup.ps1†L601-L611】【F:AllInOne-AD-Setup.ps1†L954-L967】
- Review yellow warnings in the activity log or console for ACL or CSV issues; the run continues but highlights skipped actions.【F:AllInOne-AD-Setup.ps1†L544-L558】【F:AllInOne-AD-Setup.ps1†L795-L803】

## Automated usage example
Every GUI prompt maps to a parameter so you can run the provisioning unattended when resetting labs:
```powershell
.\AllInOne-AD-Setup.ps1 -ServerName 'DC01' -BasePath 'D:\Shares' -Departments 'HR','IT','Accounting','Executives','Facilities' `
  -DefaultPassword 'P@ssw0rd!' -UPNSuffix 'lab.example.com' -CsvPath 'C:\Imports\users.csv' `
  -AltAdminDisplayName 'Lab Admin' -AltAdminSamAccountName 'labadmin' -AltAdminPassword 'Sup3rS3cret!' -SkipAclForDepartments
```
Combine with `-SkipFoldersAndShares` or `-NoGui` to reuse existing infrastructure without showing the Windows Forms interface.【F:AllInOne-AD-Setup.ps1†L19-L34】【F:AllInOne-AD-Setup.ps1†L872-L979】
