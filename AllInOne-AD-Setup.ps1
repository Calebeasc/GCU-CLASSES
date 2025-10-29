<# =======================================================================
 AllInOne-AD-Setup.ps1 (App Edition)
 -----------------------------------------------------------------------
 Provides a Windows Forms front-end for provisioning Active Directory lab
 environments. The workflow collects environment details, builds/ensures
 departmental shares and permissions, creates matching OUs and security
 groups, imports users from CSV, provisions home directories, and
 optionally creates an alternate administrative account.

 Key capabilities
   - GUI-driven configuration with defaults and validation
   - Idempotent folder/share/OU/group creation routines
   - CSV importer that accepts FullName or First/Last headers
   - Automatic home folder provisioning with ACL management
   - Domain autodetection with optional overrides
   - Rich-text run log for status visibility when using the GUI
 ======================================================================= #>

[CmdletBinding()]
param(
  # Optional parameter overrides still supported for automation scenarios.
  [string]$ServerName,
  [string]$BasePath,
  [string[]]$Departments,
  [string]$DefaultPassword,
  [string]$CsvPath,
  [string]$UPNSuffix,
  [string]$AltAdminDisplayName,
  [string]$AltAdminSamAccountName,
  [string]$AltAdminPassword,
  [switch]$SkipFoldersAndShares,
  [switch]$SkipAclForDepartments,
  [switch]$NoGui
)

# ---------------------------------------------------------------------------
# Module loading and UI initialization
# ---------------------------------------------------------------------------

Import-Module ActiveDirectory -ErrorAction Stop
Import-Module SmbShare        -ErrorAction Stop

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# Domain discovery utilities
# ---------------------------------------------------------------------------

$Domain   = Get-ADDomain
$DomainDN = $Domain.DistinguishedName
$DNSRoot  = $Domain.DNSRoot
$NetBIOS  = $Domain.NetBIOSName

$dcObj = $null
try {
  $dcObj = Get-ADDomainController -Discover -Writable -ErrorAction Stop
} catch {
  $dcObj = $null
}

$DC = $null
if ($dcObj) {
  if ($dcObj.HostName -is [string]) {
    $DC = $dcObj.HostName
  } elseif ($dcObj.HostName) {
    $firstHost = $dcObj.HostName | Select-Object -First 1
    if ($firstHost) {
      $DC = [string]$firstHost
    }
  }
}

if (-not $DC) {
  try {
    $pdc = $Domain.PdcRoleOwner
    if ($pdc -and $pdc.Name) {
      $DC = [string]$pdc.Name
    }
  } catch {
    $DC = $null
  }
}

# ---------------------------------------------------------------------------
# Logging helpers (console + optional GUI log window)
# ---------------------------------------------------------------------------

$script:LogForm    = $null
$script:LogTextBox = $null
$script:LogClose   = $null
$script:UseGuiLog  = $false

function Write-Log {
  param(
    [string]$Message,
    [ConsoleColor]$Color = [ConsoleColor]::Gray
  )

  if (-not $Message) { return }

  Write-Host $Message -ForegroundColor $Color

  if ($script:LogTextBox) {
    $script:LogTextBox.AppendText("$Message`r`n")
    $script:LogTextBox.SelectionStart = $script:LogTextBox.TextLength
    $script:LogTextBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
  }
}

function Initialize-LogWindow {
  if ($script:LogForm) { return }

  $script:LogForm = New-Object System.Windows.Forms.Form
  $script:LogForm.Text = 'AllInOne AD Setup - Activity Log'
  $script:LogForm.Size = New-Object System.Drawing.Size(820,520)
  $script:LogForm.StartPosition = 'CenterScreen'

  $script:LogTextBox = New-Object System.Windows.Forms.RichTextBox
  $script:LogTextBox.Dock = 'Fill'
  $script:LogTextBox.ReadOnly = $true
  $script:LogTextBox.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
  $script:LogTextBox.ForeColor = [System.Drawing.Color]::FromArgb(235,235,235)
  $script:LogTextBox.Font      = New-Object System.Drawing.Font('Consolas',10)

  $panel = New-Object System.Windows.Forms.Panel
  $panel.Dock = 'Bottom'
  $panel.Height = 50

  $script:LogClose = New-Object System.Windows.Forms.Button
  $script:LogClose.Text = 'Close'
  $script:LogClose.Enabled = $false
  $script:LogClose.Size = New-Object System.Drawing.Size(90,28)
  $script:LogClose.Location = New-Object System.Drawing.Point(705,10)
  $script:LogClose.Add_Click({
    $script:LogForm.Close()
  })

  $panel.Controls.Add($script:LogClose)

  $script:LogForm.Controls.Add($script:LogTextBox)
  $script:LogForm.Controls.Add($panel)

  $script:LogForm.Show()
  [System.Windows.Forms.Application]::DoEvents()
  $script:UseGuiLog = $true
}

# ---------------------------------------------------------------------------
# Path and input helpers
# ---------------------------------------------------------------------------

function Resolve-AbsolutePath {
  <#
    .SYNOPSIS
      Normalizes relative or rooted paths to absolute file system paths.
  #>
  param([Parameter(Mandatory)][string]$Path)

  if ([string]::IsNullOrWhiteSpace($Path)) {
    throw 'Path value cannot be empty when resolving to an absolute path.'
  }

  try {
    if ([System.IO.Path]::IsPathRooted($Path)) {
      return [System.IO.Path]::GetFullPath($Path)
    }

    $baseLocation = (Get-Location).ProviderPath
    $combined = Join-Path -Path $baseLocation -ChildPath $Path
    return [System.IO.Path]::GetFullPath($combined)
  } catch {
    throw "Unable to resolve path '$Path' to an absolute location: $($_.Exception.Message)"
  }
}

function ConvertTo-DepartmentList {
  <#
    .SYNOPSIS
      Splits a comma or newline separated list into a sanitized department array.
  #>
  param([string]$Text)

  if ([string]::IsNullOrWhiteSpace($Text)) {
    return @()
  }

  $parts = $Text -split '[\r\n,]'
  return $parts | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

function Show-GuiFilePicker {
  <#
    .SYNOPSIS
      Opens a file selection dialog and returns the chosen path.
  #>
  param(
    [string]$Title = 'Select a file',
    [string]$Filter = 'All files (*.*)|*.*'
  )

  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Title = $Title
  $dialog.Filter = $Filter
  $dialog.Multiselect = $false

  if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    return $dialog.FileName
  }

  return $null
}

# ---------------------------------------------------------------------------
# GUI configuration dialog
# ---------------------------------------------------------------------------

function Get-ConfigurationFromGui {
  <#
    .SYNOPSIS
      Presents a Windows Forms dialog to capture configuration settings.
  #>

  $form = New-Object System.Windows.Forms.Form
  $form.Text = 'AllInOne AD Setup - Configuration'
  $form.Size = New-Object System.Drawing.Size(720,640)
  $form.StartPosition = 'CenterScreen'
  $form.FormBorderStyle = 'FixedDialog'
  $form.MaximizeBox = $false
  $form.MinimizeBox = $false

  $font = New-Object System.Drawing.Font('Segoe UI',9)
  $form.Font = $font

  $layout = New-Object System.Windows.Forms.TableLayoutPanel
  $layout.Dock = 'Fill'
  $layout.ColumnCount = 3
  $layout.RowCount = 14
  $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,30)))
  $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,50)))
  $layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent,20)))

  for ($i = 0; $i -lt 14; $i++) {
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
  }

  $defServer = $env:COMPUTERNAME
  $defBase   = 'C:\Shares'

  # Helper for labeled textbox rows
  function Add-InputRow {
    param(
      [string]$LabelText,
      [System.Windows.Forms.Control]$Control,
      [System.Windows.Forms.Control]$Button
    )

    $row = $layout.RowCount
    $layout.RowCount++
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $LabelText
    $label.AutoSize = $true
    $label.Margin = New-Object System.Windows.Forms.Padding(6,8,6,4)

    $Control.Margin = New-Object System.Windows.Forms.Padding(6,4,6,4)
    if ($Button) {
      $Button.Margin = New-Object System.Windows.Forms.Padding(6,4,6,4)
    }

    $layout.Controls.Add($label,0,$row)
    $layout.Controls.Add($Control,1,$row)
    if ($Button) { $layout.Controls.Add($Button,2,$row) }
  }

  $serverBox = New-Object System.Windows.Forms.TextBox
  $serverBox.Text = $defServer
  Add-InputRow -LabelText 'Share host (server name)' -Control $serverBox

  $baseBox = New-Object System.Windows.Forms.TextBox
  $baseBox.Text = $defBase
  $baseBrowse = New-Object System.Windows.Forms.Button
  $baseBrowse.Text = 'Browse...'
  $baseBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Select the base directory for departmental shares'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $baseBox.Text = $dlg.SelectedPath
    }
  })
  Add-InputRow -LabelText 'Base folder for shares' -Control $baseBox -Button $baseBrowse

  $deptBox = New-Object System.Windows.Forms.TextBox
  $deptBox.Multiline = $true
  $deptBox.Height = 80
  $deptBox.ScrollBars = 'Vertical'
  $deptBox.Text = "Executives,HR,IT,Management,Accounting,Doctors,Nurses,Laboratory,Medical Records,Facilities"
  Add-InputRow -LabelText 'Departments (comma or newline separated)' -Control $deptBox

  $passwordBox = New-Object System.Windows.Forms.TextBox
  $passwordBox.UseSystemPasswordChar = $true
  $passwordBox.Text = 'Red.vine1'
  Add-InputRow -LabelText 'Default password for new users' -Control $passwordBox

  $upnBox = New-Object System.Windows.Forms.TextBox
  $upnBox.Text = $DNSRoot
  Add-InputRow -LabelText 'UPN suffix (domain)' -Control $upnBox

  $csvBox = New-Object System.Windows.Forms.TextBox
  $csvBrowse = New-Object System.Windows.Forms.Button
  $csvBrowse.Text = 'Select CSV...'
  $csvBrowse.Add_Click({
    $file = Show-GuiFilePicker -Title 'Select user CSV file'
    if ($file) { $csvBox.Text = $file }
  })
  Add-InputRow -LabelText 'User CSV path (optional)' -Control $csvBox -Button $csvBrowse

  $altDisplay = New-Object System.Windows.Forms.TextBox
  Add-InputRow -LabelText 'Alt admin display name (optional)' -Control $altDisplay

  $altSam = New-Object System.Windows.Forms.TextBox
  Add-InputRow -LabelText 'Alt admin username (sAM)' -Control $altSam

  $altPass = New-Object System.Windows.Forms.TextBox
  $altPass.UseSystemPasswordChar = $true
  Add-InputRow -LabelText 'Alt admin password' -Control $altPass

  $skipShares = New-Object System.Windows.Forms.CheckBox
  $skipShares.Text = 'Skip folder/share creation'
  $skipShares.AutoSize = $true
  $layout.Controls.Add($skipShares,1,$layout.RowCount)
  $layout.RowCount++
  $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

  $skipAcl = New-Object System.Windows.Forms.CheckBox
  $skipAcl.Text = 'Skip department ACL reset'
  $skipAcl.AutoSize = $true
  $layout.Controls.Add($skipAcl,1,$layout.RowCount)
  $layout.RowCount++
  $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))

  $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
  $buttonPanel.FlowDirection = 'RightToLeft'
  $buttonPanel.Dock = 'Bottom'
  $buttonPanel.Padding = New-Object System.Windows.Forms.Padding(6)

  $runButton = New-Object System.Windows.Forms.Button
  $runButton.Text = 'Run Provisioning'
  $runButton.Width = 150
  $runButton.Height = 32

  $cancelButton = New-Object System.Windows.Forms.Button
  $cancelButton.Text = 'Cancel'
  $cancelButton.Width = 100
  $cancelButton.Height = 32
  $cancelButton.Add_Click({
    $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Close()
  })

  $buttonPanel.Controls.Add($runButton)
  $buttonPanel.Controls.Add($cancelButton)

  $form.Controls.Add($layout)
  $form.Controls.Add($buttonPanel)

  $resultObject = $null

  $runButton.Add_Click({
    $departments = ConvertTo-DepartmentList -Text $deptBox.Text
    if ($departments.Count -lt 5) {
      [System.Windows.Forms.MessageBox]::Show('Please provide at least five department names.','Validation', 'OK', 'Error') | Out-Null
      return
    }

    $resultObject = [pscustomobject]@{
      ServerName              = if ([string]::IsNullOrWhiteSpace($serverBox.Text)) { $env:COMPUTERNAME } else { $serverBox.Text.Trim() }
      BasePath                = if ([string]::IsNullOrWhiteSpace($baseBox.Text)) { 'C:\\Shares' } else { $baseBox.Text.Trim() }
      Departments             = $departments
      DefaultPassword         = if ([string]::IsNullOrWhiteSpace($passwordBox.Text)) { 'Red.vine1' } else { $passwordBox.Text }
      CsvPath                 = if ([string]::IsNullOrWhiteSpace($csvBox.Text)) { $null } else { $csvBox.Text.Trim() }
      UPNSuffix               = if ([string]::IsNullOrWhiteSpace($upnBox.Text)) { $DNSRoot } else { $upnBox.Text.Trim() }
      AltAdminDisplayName     = if ([string]::IsNullOrWhiteSpace($altDisplay.Text)) { $null } else { $altDisplay.Text.Trim() }
      AltAdminSamAccountName  = if ([string]::IsNullOrWhiteSpace($altSam.Text)) { $null } else { $altSam.Text.Trim() }
      AltAdminPassword        = if ([string]::IsNullOrWhiteSpace($altPass.Text)) { $null } else { $altPass.Text }
      SkipFoldersAndShares    = $skipShares.Checked
      SkipAclForDepartments   = $skipAcl.Checked
    }

    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
  })

  $form.AcceptButton = $runButton
  $form.CancelButton = $cancelButton

  $dialogResult = $form.ShowDialog()
  if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
    return $null
  }

  return $resultObject
}

# ---------------------------------------------------------------------------
# Core provisioning helper functions
# ---------------------------------------------------------------------------

function Ensure-Directory {
  <#
    .SYNOPSIS
      Creates a directory when it does not already exist and logs the action.
  #>
  param([Parameter(Mandatory)][string]$Path)

  if ([string]::IsNullOrWhiteSpace($Path)) {
    throw 'Ensure-Directory was given an empty path. Please supply a valid folder location.'
  }

  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force | Out-Null
    Write-Log "Created folder: $Path" -Color Cyan
  } else {
    Write-Log "Folder already exists: $Path" -Color DarkCyan
  }
}

function Ensure-Share {
  <#
    .SYNOPSIS
      Creates or updates an SMB share with Everyone:Full access.
  #>
  param(
    [Parameter(Mandatory)][string]$Name,
    [Parameter(Mandatory)][string]$Path
  )

  if ([string]::IsNullOrWhiteSpace($Path)) {
    throw "Ensure-Share '$Name' received an empty path. Specify a valid folder to share."
  }

  $existing = Get-SmbShare -Name $Name -ErrorAction SilentlyContinue
  if (-not $existing) {
    try {
      New-SmbShare -Name $Name -Path $Path -FullAccess 'Everyone' | Out-Null
      Write-Log "Created SMB share: $Name -> $Path (Everyone: Full Control)" -Color Cyan
    } catch {
      Write-Log "FAILED to create SMB share '$Name': $($_.Exception.Message)" -Color Red
    }
  } else {
    Grant-SmbShareAccess -Name $Name -AccountName 'Everyone' -AccessRight Full -Force -ErrorAction SilentlyContinue | Out-Null
    Write-Log "Verified SMB share: $Name (Everyone: Full Control)" -Color DarkCyan
  }
}

function Ensure-OU {
  <#
    .SYNOPSIS
      Ensures an Organizational Unit exists beneath the domain root.
  #>
  param([Parameter(Mandatory)][string]$OuName)

  $ou = Get-ADOrganizationalUnit -LDAPFilter "(ou=$OuName)" -SearchBase $DomainDN -ErrorAction SilentlyContinue
  if (-not $ou) {
    New-ADOrganizationalUnit -Name $OuName -Path $DomainDN -ProtectedFromAccidentalDeletion:$false -Server $DC | Out-Null
    Write-Log "Created OU: $OuName" -Color Cyan
  } else {
    Write-Log "OU already exists: $OuName" -Color DarkCyan
  }
}

function Ensure-Group {
  <#
    .SYNOPSIS
      Ensures a global security group exists inside a specific OU.
  #>
  param(
    [Parameter(Mandatory)][string]$GroupCN,
    [Parameter(Mandatory)][string]$OuName
  )

  $ouPath = "OU=$OuName,$DomainDN"
  $grp = Get-ADGroup -LDAPFilter "(cn=$GroupCN)" -SearchBase $ouPath -ErrorAction SilentlyContinue
  if (-not $grp) {
    New-ADGroup -Name $GroupCN -SamAccountName $GroupCN -GroupCategory Security -GroupScope Global -Path $ouPath -Server $DC | Out-Null
    Write-Log "Created Group: $GroupCN (OU=$OuName)" -Color Cyan
  } else {
    Write-Log "Group already exists: $GroupCN (OU=$OuName)" -Color DarkCyan
  }
}

function New-UniqueSam {
  <#
    .SYNOPSIS
      Generates a unique sAMAccountName within the domain (<=20 characters).
  #>
  param([Parameter(Mandatory)][string]$Base)

  $b = ($Base -replace '[^A-Za-z0-9]','').ToLower()
  if ($b.Length -gt 20) { $b = $b.Substring(0,20) }
  if (-not $b) { $b = 'user' }

  $candidate = $b
  $i = 1
  while (Get-ADUser -LDAPFilter "(sAMAccountName=$candidate)" -SearchBase $DomainDN -ErrorAction SilentlyContinue) {
    $suffix = $i.ToString()
    $maxLen = 20 - $suffix.Length
    if ($maxLen -lt 1) { $maxLen = 1 }
    $candidate = if ($b.Length -gt $maxLen) { $b.Substring(0,$maxLen) + $suffix } else { $b + $suffix }
    $i++
  }

  return $candidate
}

function Get-PrincipalCandidates {
  <#
    .SYNOPSIS
      Builds candidate identifiers (SAM + SID) for icacls permission grants.
  #>
  param([Parameter(Mandatory)]$AdObject)

  $candidates = New-Object System.Collections.Generic.List[string]

  if ($AdObject -and $AdObject.SamAccountName) {
    $candidates.Add(('{0}\{1}' -f $NetBIOS, $AdObject.SamAccountName.Trim())) | Out-Null
  }

  if ($AdObject -and $AdObject.SID) {
    $sidValue = $AdObject.SID.Value
    if ($sidValue) { $candidates.Add("*$sidValue") | Out-Null }
  }

  return $candidates.ToArray()
}

function Invoke-IcaclsGrant {
  <#
    .SYNOPSIS
      Attempts to grant permissions using icacls with multiple identity options.
  #>
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string]$RuleTemplate,
    [Parameter(Mandatory)][string[]]$Candidates,
    [string]$Description
  )

  $errors = @()
  foreach ($candidate in $Candidates) {
    $rule = $RuleTemplate -f $candidate
    $result = & icacls $Path /grant $rule 2>&1
    if ($LASTEXITCODE -eq 0) {
      return $true
    }
    $errors += $result
  }

  if ($errors.Count -gt 0) {
    Write-Log "WARNING: Failed to update ACL ($Description). Tried: $([string]::Join(', ', $Candidates)). Details: $([string]::Join(' ', $errors))" -Color Yellow
  }

  return $false
}

function Invoke-IcaclsReset {
  <#
    .SYNOPSIS
      Resets inheritance and removes BUILTIN\Users from a directory.
  #>
  param([Parameter(Mandatory)][string]$Path)

  & icacls $Path /inheritance:r | Out-Null
  & icacls $Path /remove 'BUILTIN\Users' 2>$null | Out-Null
}

# ---------------------------------------------------------------------------
# Department ACL application
# ---------------------------------------------------------------------------

function Set-DepartmentAcl {
  <#
    .SYNOPSIS
      Applies the standard department ACL matrix to a folder.
  #>
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$Department,
    [Parameter(Mandatory)]$DepartmentGroup,
    [Parameter(Mandatory)]$ItGroup,
    $ExecutivesGroup,
    $ManagementGroup
  )

  Invoke-IcaclsReset -Path $Folder

  & icacls $Folder /grant 'NT AUTHORITY\SYSTEM:(OI)(CI)F' | Out-Null
  & icacls $Folder /grant 'BUILTIN\Administrators:(OI)(CI)F' | Out-Null

  $deptCandidates = Get-PrincipalCandidates -AdObject $DepartmentGroup
  Invoke-IcaclsGrant -Path $Folder -RuleTemplate '{0}:(OI)(CI)F' -Candidates $deptCandidates -Description "Department full for $Department" | Out-Null

  $itCandidates = Get-PrincipalCandidates -AdObject $ItGroup
  Invoke-IcaclsGrant -Path $Folder -RuleTemplate '{0}:(OI)(CI)F' -Candidates $itCandidates -Description 'IT full control' | Out-Null

  if ($Department -notmatch '^(?i)IT$') {
    if ($ExecutivesGroup) {
      $execCandidates = Get-PrincipalCandidates -AdObject $ExecutivesGroup
      Invoke-IcaclsGrant -Path $Folder -RuleTemplate '{0}:(OI)(CI)RX' -Candidates $execCandidates -Description 'Executives read access' | Out-Null
    }

    if ($ManagementGroup) {
      $mgmtCandidates = Get-PrincipalCandidates -AdObject $ManagementGroup
      Invoke-IcaclsGrant -Path $Folder -RuleTemplate '{0}:(OI)(CI)RX' -Candidates $mgmtCandidates -Description 'Management read access' | Out-Null
    }
  }

  Write-Log "NTFS set: $Folder" -Color DarkCyan
}

# ---------------------------------------------------------------------------
# Alternate administrator provisioning
# ---------------------------------------------------------------------------

function Ensure-AlternateAdminAccount {
  <#
    .SYNOPSIS
      Creates or updates an alternate administrative account.
  #>
  param(
    [Parameter(Mandatory)][string]$DisplayName,
    [Parameter(Mandatory)][string]$SamAccountName,
    [Parameter(Mandatory)][string]$Password,
    [Parameter(Mandatory)][string]$ItDepartmentOu,
    [Parameter(Mandatory)]$ItGroup
  )

  $secure = ConvertTo-SecureString $Password -AsPlainText -Force
  $existing = Get-ADUser -Identity $SamAccountName -ErrorAction SilentlyContinue

  if (-not $existing) {
    New-ADUser -Name $DisplayName -SamAccountName $SamAccountName -UserPrincipalName "$SamAccountName@$DNSRoot" -AccountPassword $secure -Enabled $true -PasswordNeverExpires $false -ChangePasswordAtLogon $false -Path $ItDepartmentOu -Server $DC | Out-Null
    Write-Log "Created alternate admin account: $SamAccountName" -Color Cyan
    $existing = Get-ADUser -Identity $SamAccountName
  } else {
    Write-Log "Alternate admin already exists: $SamAccountName" -Color DarkCyan
  }

  try {
    Add-ADGroupMember -Identity $ItGroup -Members $existing -Server $DC -ErrorAction SilentlyContinue
  } catch {}

  foreach ($adminGroup in @('Domain Admins','Administrators')) {
    $grp = Get-ADGroup -Identity $adminGroup -ErrorAction SilentlyContinue
    if ($grp) {
      try {
        Add-ADGroupMember -Identity $grp -Members $existing -Server $DC -ErrorAction SilentlyContinue
      } catch {}
    }
  }

  return $existing
}

# ---------------------------------------------------------------------------
# CSV user import
# ---------------------------------------------------------------------------

function Import-UsersFromCsv {
  <#
    .SYNOPSIS
      Imports users from a CSV file and provisions accounts, groups, and home folders.
  #>
  param(
    [Parameter(Mandatory)][string]$CsvPath,
    [Parameter(Mandatory)][string]$DefaultPassword,
    [Parameter(Mandatory)][string]$ServerName,
    [Parameter(Mandatory)][string]$BasePath,
    [Parameter(Mandatory)][string[]]$Departments,
    [Parameter(Mandatory)]$DeptGroupsMap,
    [Parameter(Mandatory)]$ItGroup,
    $ExecutivesGroup,
    $ManagementGroup,
    [Parameter(Mandatory)][string]$UpnSuffix
  )

  $result = [pscustomobject]@{
    Created = 0
    Updated = 0
    Malformed = New-Object System.Collections.Generic.List[string]
  }

  $securePwd = ConvertTo-SecureString $DefaultPassword -AsPlainText -Force

  $rows = Import-Csv -LiteralPath $CsvPath
  if (-not $rows -or $rows.Count -eq 0) {
    throw "CSV is empty: $CsvPath"
  }

  $headers = $rows[0].psobject.Properties.Name
  $fullHdr  = ($headers | Where-Object { $_ -match '^(FullName|Name)$' } | Select-Object -First 1)
  $firstHdr = ($headers | Where-Object { $_ -match '^(FirstName|First Name|GivenName|Given Name|FName)$' } | Select-Object -First 1)
  $lastHdr  = ($headers | Where-Object { $_ -match '^(LastName|Last Name|Surname|Sur Name|LName)$' }  | Select-Object -First 1)

  if (-not $fullHdr -and (-not $firstHdr -or -not $lastHdr)) {
    throw 'CSV must have FullName OR First/Last columns.'
  }

  Write-Log "Importing users from $CsvPath ..." -Color Yellow

  foreach ($row in $rows) {
    $First = $null
    $Last  = $null

    if ($fullHdr) {
      $raw = ([string]$row.$fullHdr).Trim()
      if (-not $raw) { continue }
      $parts = $raw -split '\s+'
      if ($parts.Count -lt 2) {
        $result.Malformed.Add($raw) | Out-Null
        continue
      }
      $First = $parts[0]
      $Last  = $parts[-1]
    } else {
      $First = [string]$row.$firstHdr
      $Last  = [string]$row.$lastHdr
      if ([string]::IsNullOrWhiteSpace($First) -or [string]::IsNullOrWhiteSpace($Last)) {
        $result.Malformed.Add("$First $Last") | Out-Null
        continue
      }
      $First = $First.Trim()
      $Last  = $Last.Trim()
    }

    $baseSam = ($First.Substring(0,1) + $Last)
    $sam = New-UniqueSam -Base $baseSam
    $upn = "$sam@$UpnSuffix"

    $dept = Get-Random -InputObject $Departments
    $ouPath = "OU=$dept,$DomainDN"
    $deptGroup = $DeptGroupsMap[$dept]

    $userObj = Get-ADUser -Identity $sam -ErrorAction SilentlyContinue

    if (-not $userObj) {
      try {
        New-ADUser -Name "$First $Last" -GivenName $First -Surname $Last -SamAccountName $sam -UserPrincipalName $upn -Path $ouPath -AccountPassword $securePwd -Enabled $true -ChangePasswordAtLogon $true -Server $DC | Out-Null
        $result.Created++
        Write-Log "Created user: $sam ($First $Last) -> $dept" -Color Green
      } catch {
        Write-Log "Failed to create user $First $Last: $($_.Exception.Message)" -Color Red
        $result.Malformed.Add("$First $Last (creation failed)") | Out-Null
        continue
      }
      $userObj = Get-ADUser -Identity $sam -ErrorAction SilentlyContinue
    } else {
      $result.Updated++
      Write-Log "User already existed, updating: $sam" -Color DarkGreen
    }

    if ($userObj) {
      try {
        Move-ADObject -Identity $userObj.DistinguishedName -TargetPath $ouPath -Server $DC -ErrorAction SilentlyContinue
      } catch {}

      if ($deptGroup) {
        try { Add-ADGroupMember -Identity $deptGroup -Members $userObj -Server $DC -ErrorAction SilentlyContinue } catch {}
      }
      if ($ExecutivesGroup -and $dept -ieq 'Executives') {
        try { Add-ADGroupMember -Identity $ExecutivesGroup -Members $userObj -Server $DC -ErrorAction SilentlyContinue } catch {}
      }
      if ($ManagementGroup -and $dept -ieq 'Management') {
        try { Add-ADGroupMember -Identity $ManagementGroup -Members $userObj -Server $DC -ErrorAction SilentlyContinue } catch {}
      }

      try {
        Set-ADUser -Identity $userObj -HomeDrive 'H:' -HomeDirectory "\\$ServerName\Home\$sam" -Server $DC -ErrorAction SilentlyContinue
      } catch {}

      $homeLocal = Join-Path -Path $BasePath -ChildPath (Join-Path -Path 'Home' -ChildPath $sam)
      if (-not (Test-Path -LiteralPath $homeLocal)) {
        New-Item -ItemType Directory -Path $homeLocal -Force | Out-Null
        Write-Log "Created home folder: $homeLocal" -Color Cyan
      } else {
        Write-Log "Ensured home folder: $homeLocal" -Color DarkCyan
      }

      Invoke-IcaclsReset -Path $homeLocal
      & icacls $homeLocal /grant 'NT AUTHORITY\SYSTEM:(OI)(CI)F' | Out-Null
      & icacls $homeLocal /grant 'BUILTIN\Administrators:(OI)(CI)F' | Out-Null

      $userCandidates = Get-PrincipalCandidates -AdObject $userObj
      Invoke-IcaclsGrant -Path $homeLocal -RuleTemplate '{0}:(OI)(CI)F' -Candidates $userCandidates -Description "Home full for $sam" | Out-Null
      $itCandidates = Get-PrincipalCandidates -AdObject $ItGroup
      Invoke-IcaclsGrant -Path $homeLocal -RuleTemplate '{0}:(OI)(CI)F' -Candidates $itCandidates -Description 'IT home access' | Out-Null
    }
  }

  if ($result.Malformed.Count -gt 0) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $outPath = Join-Path -Path $desktop -ChildPath 'MalformedNames.txt'
    $result.Malformed | Set-Content -Path $outPath -Encoding UTF8
    Write-Log "Some rows were skipped. See: $outPath" -Color Yellow
    $result | Add-Member -NotePropertyName MalformedPath -NotePropertyValue $outPath -Force
  }

  return $result
}

# ---------------------------------------------------------------------------
# Summary output helper
# ---------------------------------------------------------------------------

function Show-Verification {
  <#
    .SYNOPSIS
      Displays post-run verification information to the log window/console.
  #>
  param(
    [Parameter(Mandatory)][string[]]$Departments,
    [Parameter(Mandatory)][string]$BasePath,
    [Parameter(Mandatory)]$DeptGroupsMap,
    $ImportResult,
    $AltAdminAccount,
    [Parameter(Mandatory)][string]$ServerName
  )

  Write-Log ''
  Write-Log '--- Verification Summary ---' -Color Cyan

  Write-Log "Shares on $ServerName:" -Color Cyan
  (Get-SmbShare | Where-Object { $_.Name -in ($Departments + 'Home' + 'Profiles') }) | ForEach-Object {
    Write-Log ("  {0,-20} {1}" -f $_.Name, $_.Path)
  }

  Write-Log ''
  Write-Log 'Organizational Units:' -Color Cyan
  $ouFilterParts = $Departments | ForEach-Object { "(ou=$_)" }
  $ouFilter = "(|" + ($ouFilterParts -join '') + ")"
  Get-ADOrganizationalUnit -SearchBase $DomainDN -LDAPFilter $ouFilter | Sort-Object Name | ForEach-Object {
    Write-Log ("  {0}" -f $_.Name)
  }

  Write-Log ''
  Write-Log 'Groups:' -Color Cyan
  foreach ($dept in $Departments) {
    $grp = $DeptGroupsMap[$dept]
    if ($grp) {
      Write-Log ("  {0,-20} {1}" -f $dept, $grp.Name)
    } else {
      Write-Log ("  {0,-20} <missing>" -f $dept)
    }
  }

  if ($ImportResult) {
    Write-Log ''
    Write-Log ("Users created: {0}, updated: {1}" -f $ImportResult.Created, $ImportResult.Updated) -Color Green
    if ($ImportResult.PSObject.Properties['MalformedPath']) {
      Write-Log "Malformed names log: $($ImportResult.MalformedPath)" -Color Yellow
    }
  }

  if ($AltAdminAccount) {
    Write-Log ''
    Write-Log ("Alternate admin: {0}" -f $AltAdminAccount.SamAccountName) -Color Cyan
  }

  Write-Log ''
  Write-Log 'Done.' -Color Green
}

# ---------------------------------------------------------------------------
# Main execution logic
# ---------------------------------------------------------------------------

$configuration = $null

if (-not $NoGui -and $PSBoundParameters.Count -eq 0) {
  $configuration = Get-ConfigurationFromGui
  if (-not $configuration) {
    Write-Host 'Operation cancelled by user.' -ForegroundColor Yellow
    return
  }
  Initialize-LogWindow
} else {
  $configuration = [pscustomobject]@{
    ServerName             = if ([string]::IsNullOrWhiteSpace($ServerName)) { $env:COMPUTERNAME } else { $ServerName.Trim() }
    BasePath               = if ([string]::IsNullOrWhiteSpace($BasePath)) { 'C:\\Shares' } else { $BasePath.Trim() }
    Departments            = if ($Departments) { $Departments } else { @('Executives','HR','IT','Management','Accounting','Doctors','Nurses','Laboratory','Medical Records','Facilities') }
    DefaultPassword        = if ([string]::IsNullOrWhiteSpace($DefaultPassword)) { 'Red.vine1' } else { $DefaultPassword }
    CsvPath                = if ([string]::IsNullOrWhiteSpace($CsvPath)) { $null } else { $CsvPath }
    UPNSuffix              = if ([string]::IsNullOrWhiteSpace($UPNSuffix)) { $DNSRoot } else { $UPNSuffix.Trim() }
    AltAdminDisplayName    = if ([string]::IsNullOrWhiteSpace($AltAdminDisplayName)) { $null } else { $AltAdminDisplayName }
    AltAdminSamAccountName = if ([string]::IsNullOrWhiteSpace($AltAdminSamAccountName)) { $null } else { $AltAdminSamAccountName }
    AltAdminPassword       = if ([string]::IsNullOrWhiteSpace($AltAdminPassword)) { $null } else { $AltAdminPassword }
    SkipFoldersAndShares   = [bool]$SkipFoldersAndShares
    SkipAclForDepartments  = [bool]$SkipAclForDepartments
  }
}

if (-not $configuration) {
  Write-Host 'Configuration could not be determined.' -ForegroundColor Red
  return
}

if ([string]::IsNullOrWhiteSpace($configuration.BasePath)) {
  $configuration.BasePath = 'C:\Shares'
}

try {
  $configuration.BasePath = Resolve-AbsolutePath -Path $configuration.BasePath
} catch {
  Write-Host $_.Exception.Message -ForegroundColor Red
  return
}

if ($configuration.CsvPath) {
  if (-not (Test-Path -LiteralPath $configuration.CsvPath)) {
    Write-Host "CSV not found: $($configuration.CsvPath)" -ForegroundColor Red
    return
  }
}

$departments = $configuration.Departments
if ($departments.Count -lt 5) {
  Write-Host 'Please specify at least five departments.' -ForegroundColor Red
  return
}

Write-Log "Domain: $DNSRoot  |  DN: $DomainDN  |  NetBIOS: $NetBIOS" -Color Cyan
Write-Log "Server for shares: $($configuration.ServerName)" -Color Cyan
Write-Log "Base path: $($configuration.BasePath)" -Color Cyan
Write-Log "Departments: $([string]::Join(', ', $departments))" -Color Cyan

$deptGroups = @{}
$executivesGroup = $null
$managementGroup = $null
$itGroup = $null

if (-not $configuration.SkipFoldersAndShares) {
  Ensure-Directory -Path $configuration.BasePath
  Ensure-Directory -Path (Join-Path -Path $configuration.BasePath -ChildPath 'Home')
  Ensure-Directory -Path (Join-Path -Path $configuration.BasePath -ChildPath 'Profiles')
  Ensure-Share -Name 'Home' -Path (Join-Path -Path $configuration.BasePath -ChildPath 'Home')
  Ensure-Share -Name 'Profiles' -Path (Join-Path -Path $configuration.BasePath -ChildPath 'Profiles')
}

foreach ($dept in $departments) {
  $folderPath = Join-Path -Path $configuration.BasePath -ChildPath $dept
  if (-not $configuration.SkipFoldersAndShares) {
    Ensure-Directory -Path $folderPath
    Ensure-Share -Name $dept -Path $folderPath
  }
  Ensure-OU -OuName $dept
  Ensure-Group -GroupCN $dept -OuName $dept
  $deptGroups[$dept] = Get-ADGroup -Identity $dept -ErrorAction SilentlyContinue
  if ($dept -ieq 'Executives') { $executivesGroup = $deptGroups[$dept] }
  if ($dept -ieq 'Management') { $managementGroup = $deptGroups[$dept] }
  if ($dept -ieq 'IT') { $itGroup = $deptGroups[$dept] }
}

if (-not $itGroup) {
  Write-Log 'IT group could not be resolved; some permissions may fail.' -Color Yellow
}

if (-not $configuration.SkipAclForDepartments -and $itGroup) {
  foreach ($dept in $departments) {
    $folderPath = Join-Path -Path $configuration.BasePath -ChildPath $dept
    $deptGroup = $deptGroups[$dept]
    if (-not $deptGroup) {
      Write-Log "Skipping ACL for $dept because the group was not found." -Color Yellow
      continue
    }
    Set-DepartmentAcl -Folder $folderPath -Department $dept -DepartmentGroup $deptGroup -ItGroup $itGroup -ExecutivesGroup $executivesGroup -ManagementGroup $managementGroup
  }
}

$altAdminAccount = $null
if ($configuration.AltAdminDisplayName -and $configuration.AltAdminSamAccountName -and $configuration.AltAdminPassword -and $itGroup) {
  $itOuPath = "OU=IT,$DomainDN"
  $altAdminAccount = Ensure-AlternateAdminAccount -DisplayName $configuration.AltAdminDisplayName -SamAccountName $configuration.AltAdminSamAccountName -Password $configuration.AltAdminPassword -ItDepartmentOu $itOuPath -ItGroup $itGroup
}

$importResult = $null
if ($configuration.CsvPath) {
  $importResult = Import-UsersFromCsv -CsvPath $configuration.CsvPath -DefaultPassword $configuration.DefaultPassword -ServerName $configuration.ServerName -BasePath $configuration.BasePath -Departments $departments -DeptGroupsMap $deptGroups -ItGroup $itGroup -ExecutivesGroup $executivesGroup -ManagementGroup $managementGroup -UpnSuffix $configuration.UPNSuffix
}

Show-Verification -Departments $departments -BasePath $configuration.BasePath -DeptGroupsMap $deptGroups -ImportResult $importResult -AltAdminAccount $altAdminAccount -ServerName $configuration.ServerName

if ($script:LogClose) {
  $script:LogClose.Enabled = $true
  [System.Windows.Forms.MessageBox]::Show('Provisioning complete. Review the log for details and press Close when finished.','AllInOne AD Setup') | Out-Null
}

