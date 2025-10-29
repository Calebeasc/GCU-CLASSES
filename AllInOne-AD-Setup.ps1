<# =======================================================================
 AllInOne-AD-Setup.ps1
 - Prompts for environment-specific values (server, base path, departments, passwords)
 - Creates or verifies C:\Shares (dept folders + Home + Profiles) + SMB shares
 - Creates matching Organizational Units and Global Security groups
 - Applies department NTFS defaults (Dept & IT = Full; Exec/Mgmt = RX, except IT)
 - Optionally imports users from CSV, sets H:, creates per-user home folders + ACLs
 ======================================================================= #>

[CmdletBinding()]
param(
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
  [switch]$SkipAclForDepartments
)

# Load required modules for Active Directory and SMB management.
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module SmbShare        -ErrorAction Stop

# Load Windows Forms assemblies so GUI prompts can be displayed in both PowerShell ISE and console.
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# GUI helper utilities
# ---------------------------------------------------------------------------

function Show-GuiInputDialog {
  <#
    .SYNOPSIS
      Displays a simple dialog with a prompt, textbox, and OK/Cancel buttons.
  #>
  param(
    [Parameter(Mandatory)][string]$Title,
    [Parameter(Mandatory)][string]$Prompt,
    [string]$DefaultText
  )

  $form = New-Object System.Windows.Forms.Form
  $form.Text = $Title
  $form.StartPosition = 'CenterScreen'
  $form.FormBorderStyle = 'FixedDialog'
  $form.MaximizeBox = $false
  $form.MinimizeBox = $false
  $form.Size = New-Object System.Drawing.Size(420,180)

  $label = New-Object System.Windows.Forms.Label
  $label.Text = $Prompt
  $label.AutoSize = $true
  $label.Location = New-Object System.Drawing.Point(12,12)
  $label.MaximumSize = New-Object System.Drawing.Size(380,0)
  $form.Controls.Add($label)

  $textBox = New-Object System.Windows.Forms.TextBox
  $textBox.Size = New-Object System.Drawing.Size(380,20)
  $textBox.Location = New-Object System.Drawing.Point(12,70)
  $textBox.Text = $DefaultText
  $form.Controls.Add($textBox)

  $okButton = New-Object System.Windows.Forms.Button
  $okButton.Text = 'OK'
  $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
  $okButton.Location = New-Object System.Drawing.Point(220,110)
  $form.Controls.Add($okButton)

  $cancelButton = New-Object System.Windows.Forms.Button
  $cancelButton.Text = 'Cancel'
  $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
  $cancelButton.Location = New-Object System.Drawing.Point(310,110)
  $form.Controls.Add($cancelButton)

  $form.AcceptButton = $okButton
  $form.CancelButton = $cancelButton

  $dialogResult = $form.ShowDialog()
  if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    return $textBox.Text
  }

  return $null
}

function Show-GuiFilePicker {
  <#
    .SYNOPSIS
      Displays a file selection dialog and returns the chosen file path.
  #>
  param(
    [Parameter(Mandatory)][string]$Title,
    [string]$Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
  )

  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Title = $Title
  $dialog.Filter = $Filter
  $dialog.Multiselect = $false

  $result = $dialog.ShowDialog()
  if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    return $dialog.FileName
  }

  return $null
}

# Collect core domain metadata for naming and LDAP operations.
$Domain   = Get-ADDomain
$DomainDN = $Domain.DistinguishedName
$DNSRoot  = $Domain.DNSRoot
$NetBIOS  = $Domain.NetBIOSName

# Discover a writable domain controller to standardize subsequent AD operations.
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
# Path helper to coerce relative inputs into absolute file system paths.
# ---------------------------------------------------------------------------

function Resolve-AbsolutePath {
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

# ---------------------------------------------------------------------------
# Gather runtime inputs (GUI prompts when parameters are not supplied)
# ---------------------------------------------------------------------------

$defaultServerName = $env:COMPUTERNAME
if (-not $PSBoundParameters.ContainsKey('ServerName') -or [string]::IsNullOrWhiteSpace($ServerName)) {
  $serverPrompt = Show-GuiInputDialog -Title 'Share Host' -Prompt ("Enter the server name that will host shares. Leave blank to use {0}." -f $defaultServerName) -DefaultText $defaultServerName
  if ([string]::IsNullOrWhiteSpace($serverPrompt)) {
    $ServerName = $defaultServerName
  } else {
    $ServerName = $serverPrompt.Trim()
  }
} else {
  $ServerName = $ServerName.Trim()
}

$defaultBasePath = 'C:\Shares'
if (-not $PSBoundParameters.ContainsKey('BasePath') -or [string]::IsNullOrWhiteSpace($BasePath)) {
  $basePrompt = Show-GuiInputDialog -Title 'Base Path' -Prompt ("Enter base folder for shares. Leave blank to use {0}." -f $defaultBasePath) -DefaultText $defaultBasePath
  if ([string]::IsNullOrWhiteSpace($basePrompt)) {
    $BasePath = $defaultBasePath
  } else {
    $BasePath = $basePrompt.Trim()
  }
} else {
  $BasePath = $BasePath.Trim()
}

$BasePath = Resolve-AbsolutePath -Path $BasePath

$defaultDeptString = 'Executives, HR, IT, Management, Accounting, Doctors, Nurses, Laboratory, Medical Records, Facilities'
if (-not $PSBoundParameters.ContainsKey('Departments') -or -not $Departments -or $Departments.Count -eq 0) {
  do {
    $deptPrompt = Show-GuiInputDialog -Title 'Departments' -Prompt 'Enter ~10 department names, comma-separated (e.g., HR, IT, Accounting, ...). Leave blank to use the suggested defaults.' -DefaultText $defaultDeptString
    if ([string]::IsNullOrWhiteSpace($deptPrompt)) {
      $deptPrompt = $defaultDeptString
    }
    $Departments = $deptPrompt -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if ($Departments.Count -lt 5) {
      [void][System.Windows.Forms.MessageBox]::Show('Please enter at least 5 departments.','Validation',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)
    }
  } while ($Departments.Count -lt 5)
} else {
  $Departments = $Departments | ForEach-Object { $_.Trim() } | Where-Object { $_ }
}

if ($Departments.Count -lt 5) {
  throw 'Please enter at least 5 departments.'
}

$defaultPasswordValue = 'Red.vine1'
if (-not $PSBoundParameters.ContainsKey('DefaultPassword') -or [string]::IsNullOrWhiteSpace($DefaultPassword)) {
  $pwdPrompt = Show-GuiInputDialog -Title 'Default Password' -Prompt ("Enter default password for new users. Leave blank to use {0}." -f $defaultPasswordValue) -DefaultText $defaultPasswordValue
  if ([string]::IsNullOrWhiteSpace($pwdPrompt)) {
    $DefaultPassword = $defaultPasswordValue
  } else {
    $DefaultPassword = $pwdPrompt.Trim()
  }
} else {
  $DefaultPassword = $DefaultPassword.Trim()
}

$SecurePwd = ConvertTo-SecureString $DefaultPassword -AsPlainText -Force

$altAdminDefaultName = 'Lab Administrator'
if (-not $PSBoundParameters.ContainsKey('AltAdminDisplayName') -or [string]::IsNullOrWhiteSpace($AltAdminDisplayName)) {
  $altAdminNamePrompt = Show-GuiInputDialog -Title 'Alternate Admin Name' -Prompt 'Enter the display name for the alternate administrative account.' -DefaultText $altAdminDefaultName
  if ([string]::IsNullOrWhiteSpace($altAdminNamePrompt)) {
    $AltAdminDisplayName = $altAdminDefaultName
  } else {
    $AltAdminDisplayName = $altAdminNamePrompt.Trim()
  }
} else {
  $AltAdminDisplayName = $AltAdminDisplayName.Trim()
}

if ([string]::IsNullOrWhiteSpace($AltAdminDisplayName)) {
  throw 'Alternate admin display name cannot be blank.'
}

$altAdminDefaultSam = 'labadmin'
if (-not $PSBoundParameters.ContainsKey('AltAdminSamAccountName') -or [string]::IsNullOrWhiteSpace($AltAdminSamAccountName)) {
  $altAdminSamPrompt = Show-GuiInputDialog -Title 'Alternate Admin Username' -Prompt 'Enter the desired sAMAccountName for the alternate administrative account.' -DefaultText $altAdminDefaultSam
  if ([string]::IsNullOrWhiteSpace($altAdminSamPrompt)) {
    $AltAdminSamAccountName = $altAdminDefaultSam
  } else {
    $AltAdminSamAccountName = $altAdminSamPrompt.Trim()
  }
} else {
  $AltAdminSamAccountName = $AltAdminSamAccountName.Trim()
}

if ([string]::IsNullOrWhiteSpace($AltAdminSamAccountName)) {
  $AltAdminSamAccountName = $altAdminDefaultSam
}

$altAdminPasswordDefault = $DefaultPassword
if (-not $PSBoundParameters.ContainsKey('AltAdminPassword') -or [string]::IsNullOrWhiteSpace($AltAdminPassword)) {
  $altAdminPasswordPrompt = Show-GuiInputDialog -Title 'Alternate Admin Password' -Prompt ('Enter the password for the alternate administrative account. Leave blank to reuse {0}.' -f $altAdminPasswordDefault) -DefaultText $altAdminPasswordDefault
  if ([string]::IsNullOrWhiteSpace($altAdminPasswordPrompt)) {
    $AltAdminPassword = $altAdminPasswordDefault
  } else {
    $AltAdminPassword = $altAdminPasswordPrompt.Trim()
  }
} else {
  $AltAdminPassword = $AltAdminPassword.Trim()
}

if ([string]::IsNullOrWhiteSpace($AltAdminPassword)) {
  throw 'Alternate admin password cannot be blank.'
}

$AltAdminSecurePwd = ConvertTo-SecureString $AltAdminPassword -AsPlainText -Force

if (-not $PSBoundParameters.ContainsKey('UPNSuffix') -or [string]::IsNullOrWhiteSpace($UPNSuffix)) {
  $upnPrompt = Show-GuiInputDialog -Title 'UPN Suffix' -Prompt ("UPN domain suffix. Press OK to use {0}." -f $DNSRoot) -DefaultText $DNSRoot
  if ([string]::IsNullOrWhiteSpace($upnPrompt)) {
    $UPNSuffix = $DNSRoot
  } else {
    $UPNSuffix = $upnPrompt.Trim()
  }
} else {
  $UPNSuffix = $UPNSuffix.Trim()
}

if (-not $PSBoundParameters.ContainsKey('CsvPath')) {
  $importDecision = [System.Windows.Forms.MessageBox]::Show('Would you like to import users from a CSV file now?','CSV Import',[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
  if ($importDecision -eq [System.Windows.Forms.DialogResult]::Yes) {
    $selectedCsv = Show-GuiFilePicker -Title 'Select CSV file for user import'
    while ($selectedCsv -and -not (Test-Path -LiteralPath $selectedCsv)) {
      [void][System.Windows.Forms.MessageBox]::Show(('CSV not found: {0}' -f $selectedCsv),'CSV Import',[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error)
      $selectedCsv = Show-GuiFilePicker -Title 'Select CSV file for user import'
    }
    if ($selectedCsv) {
      $CsvPath = $selectedCsv
    }
  }
} elseif ($CsvPath) {
  $CsvPath = $CsvPath.Trim()
}

if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
  if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw ("CSV not found: {0}" -f $CsvPath)
  }
} else {
  $CsvPath = $null
}

# Summarize the runtime context so the operator can confirm the selections.
Write-Host "Domain: $DNSRoot  |  DN: $DomainDN  |  NetBIOS: $NetBIOS" -ForegroundColor Cyan
Write-Host "Server for shares: $ServerName" -ForegroundColor Cyan
Write-Host "Base path: $BasePath" -ForegroundColor Cyan
Write-Host "Departments: $($Departments -join ', ')" -ForegroundColor Cyan
Write-Host

$adServerParams = @{}
if (-not [string]::IsNullOrWhiteSpace($DC)) {
  $adServerParams['Server'] = $DC
}

# Cache any custom department -> group mappings to keep naming flexible.
$DeptNameMap = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)

function Resolve-DeptGroupName {
  param([Parameter(Mandatory)][string]$Department)
  if ($DeptNameMap.ContainsKey($Department)) {
    return $DeptNameMap[$Department]
  }
  return $Department
}

# Ensure the provided directory exists, creating it when missing.
function Ensure-Directory {
  param([Parameter(Mandatory)][string]$Path)

  $resolvedPath = Resolve-AbsolutePath -Path $Path

  if (-not (Test-Path -LiteralPath $resolvedPath)) {
    New-Item -ItemType Directory -Path $resolvedPath -Force | Out-Null
    Write-Host ("Created folder: {0}" -f $resolvedPath) -ForegroundColor Green
  } else {
    Write-Host ("Folder already exists: {0}" -f $resolvedPath) -ForegroundColor Cyan
  }
}

# Run icacls with consistent logging and error capture.
function Invoke-IcaclsCommand {
  param(
    [Parameter(Mandatory)][string[]]$Arguments,
    [string]$Description,
    [switch]$Quiet
  )

  $global:LASTEXITCODE = 0
  $output = & icacls @Arguments 2>&1
  $success = $LASTEXITCODE -eq 0
  if (-not $success -and -not $Quiet) {
    $message = if ($Description) { "Failed to update ACL ($Description): $($output -join ' ')" } else { "Failed to update ACL: $($output -join ' ')" }
    Write-Warning $message
  }

  return [pscustomobject]@{
    Success = $success
    Output  = $output
  }
}

# Build a list of SAM names and SID fallbacks for ACL operations.
function Get-IcaclsPrincipalCandidates {
  param(
    [string]$SamAccountName,
    [object]$SidValue
  )

  $candidates = New-Object System.Collections.Generic.List[string]

  if (-not [string]::IsNullOrWhiteSpace($SamAccountName)) {
    $candidates.Add(('{0}\{1}' -f $NetBIOS,$SamAccountName.Trim())) | Out-Null
  }

  if ($SidValue) {
    $sidObjects = @()

    if ($SidValue -is [System.Collections.IEnumerable] -and -not ($SidValue -is [string])) {
      foreach ($sidEntry in $SidValue) {
        if ($sidEntry) { $sidObjects += $sidEntry }
      }
    } else {
      $sidObjects = @($SidValue)
    }

    foreach ($sidObj in $sidObjects) {
      if (-not $sidObj) { continue }

      $sidString = $null
      if ($sidObj -is [System.Security.Principal.SecurityIdentifier]) {
        $sidString = $sidObj.Value
      } else {
        $sidString = [string]$sidObj
      }

      if (-not [string]::IsNullOrWhiteSpace($sidString)) {
        $sidString = $sidString.Trim()
        $sidParts = $sidString -split '\s+' | Where-Object { $_ }
        foreach ($sidPart in $sidParts) {
          $candidates.Add(('*{0}' -f $sidPart)) | Out-Null
        }
      }
    }
  }

  return $candidates.ToArray()
}

# Attempt to grant an ACL entry using each candidate principal in turn.
function Grant-IcaclsPermission {
  param(
    [Parameter(Mandatory)][string]$Path,
    [Parameter(Mandatory)][string[]]$PrincipalCandidates,
    [Parameter(Mandatory)][string]$AccessRule,
    [string]$Description
  )

  if (-not $PrincipalCandidates -or $PrincipalCandidates.Count -eq 0) {
    return $false
  }

  $lastFailure = $null

  foreach ($principal in $PrincipalCandidates) {
    if ([string]::IsNullOrWhiteSpace($principal)) { continue }
    $rule = '{0}:{1}' -f $principal,$AccessRule
    $result = Invoke-IcaclsCommand -Arguments @($Path,'/grant',$rule) -Description $Description -Quiet
    if ($result.Success) {
      return $true
    }
    $lastFailure = $result
  }

  if ($lastFailure) {
    $principalList = ($PrincipalCandidates | Where-Object { $_ }) -join ', '
    $details = if ($lastFailure.Output) { $lastFailure.Output -join ' ' } else { 'No output from icacls.' }
    if ($Description) {
      Write-Warning ("Failed to update ACL ({0}). Tried: {1}. Details: {2}" -f $Description,$principalList,$details)
    } else {
      Write-Warning ("Failed to update ACL. Tried: {0}. Details: {1}" -f $principalList,$details)
    }
  }

  return $false
}

# Ensure the SMB share exists with Everyone:Full share permissions.
function Ensure-Share {
  param(
    [Parameter(Mandatory)][string]$Name,
    [Parameter(Mandatory)][string]$Path
  )

  $resolvedPath = Resolve-AbsolutePath -Path $Path
  $existing = Get-SmbShare -Name $Name -ErrorAction SilentlyContinue
  if (-not $existing) {
    try {
      New-SmbShare -Name $Name -Path $resolvedPath -FullAccess 'Everyone' -ErrorAction Stop | Out-Null
      Write-Host ("Created SMB share: {0} -> {1} (Everyone: Full Control)" -f $Name,$resolvedPath) -ForegroundColor Green
    } catch {
      $postCheck = Get-SmbShare -Name $Name -ErrorAction SilentlyContinue
      if ($postCheck) {
        Write-Host ("SMB share already exists: {0} (Everyone: Full Control verified)" -f $Name) -ForegroundColor Cyan
      } else {
        Write-Warning ("Failed to create SMB share '{0}': {1}" -f $Name,$_.Exception.Message)
      }
    }
  } else {
    Grant-SmbShareAccess -Name $Name -AccountName 'Everyone' -AccessRight Full -Force -ErrorAction SilentlyContinue | Out-Null
    Write-Host ("SMB share already exists: {0} (Everyone: Full Control verified)" -f $Name) -ForegroundColor Cyan
  }
}

# Ensure the Organizational Unit exists (idempotent creation).
function Ensure-OU {
  param([Parameter(Mandatory)][string]$OuName)
  $ouLookupParams = @{ LDAPFilter = "(ou=$OuName)"; SearchBase = $DomainDN; ErrorAction = 'SilentlyContinue' }
  $ouLookupParams += $adServerParams
  $ou = Get-ADOrganizationalUnit @ouLookupParams
  if (-not $ou) {
    try {
      $newOuParams = @{ Name = $OuName; Path = $DomainDN; ProtectedFromAccidentalDeletion = $false }
      $newOuParams += $adServerParams
      New-ADOrganizationalUnit @newOuParams | Out-Null
      Write-Host ("Created OU: {0}" -f $OuName) -ForegroundColor Green
    } catch {
      $ou = Get-ADOrganizationalUnit @ouLookupParams
      if ($ou) {
        Write-Host ("OU already exists: {0}" -f $OuName) -ForegroundColor Cyan
      } else {
        Write-Warning ("Failed to create OU '{0}': {1}" -f $OuName,$_.Exception.Message)
      }
    }
  } else {
    Write-Host ("OU already exists: {0}" -f $OuName) -ForegroundColor Cyan
  }
}

# Ensure the departmental global security group exists under the OU.
function Ensure-Group {
  param(
    [Parameter(Mandatory)][string]$GroupCN,
    [Parameter(Mandatory)][string]$OuName
  )
  $ouPath = "OU=$OuName,$DomainDN"
  $grpLookupParams = @{ LDAPFilter = "(cn=$GroupCN)"; SearchBase = $ouPath; ErrorAction = 'SilentlyContinue' }
  $grpLookupParams += $adServerParams
  $grp = Get-ADGroup @grpLookupParams
  if (-not $grp) {
    try {
      $newGroupParams = @{ Name = $GroupCN; SamAccountName = $GroupCN; GroupCategory = 'Security'; GroupScope = 'Global'; Path = $ouPath }
      $newGroupParams += $adServerParams
      New-ADGroup @newGroupParams | Out-Null
      Write-Host ("Created Group: {0}  (in OU={1})" -f $GroupCN,$OuName) -ForegroundColor Green
    } catch {
      $grp = Get-ADGroup @grpLookupParams
      if ($grp) {
        Write-Host ("Group already exists: {0}  (in OU={1})" -f $GroupCN,$OuName) -ForegroundColor Cyan
      } else {
        Write-Warning ("Failed to create group '{0}' in OU '{1}': {2}" -f $GroupCN,$OuName,$_.Exception.Message)
      }
    }
  } else {
    Write-Host ("Group already exists: {0}  (in OU={1})" -f $GroupCN,$OuName) -ForegroundColor Cyan
  }
}

# Produce a unique sAMAccountName (<=20 chars) within the domain.
function New-UniqueSam {
  param([Parameter(Mandatory)][string]$Base)
  $b = ($Base -replace '[^A-Za-z0-9]','').ToLower()
  if ($b.Length -gt 20) {
    $b = $b.Substring(0,20)
  }
  $candidate = $b
  $i = 1
  $userLookupParams = @{ LDAPFilter = "(sAMAccountName=$candidate)"; SearchBase = $DomainDN; ErrorAction = 'SilentlyContinue' }
  $userLookupParams += $adServerParams
  while (Get-ADUser @userLookupParams) {
    $suffix = $i.ToString()
    $maxBase = 20 - $suffix.Length
    if ($b.Length -gt $maxBase) {
      $candidate = $b.Substring(0,$maxBase) + $suffix
    } else {
      $candidate = $b + $suffix
    }
    $i++
    $userLookupParams['LDAPFilter'] = "(sAMAccountName=$candidate)"
  }
  return $candidate
}

function Ensure-AlternateAdminAccount {
  param(
    [Parameter(Mandatory)][string]$DisplayName,
    [Parameter(Mandatory)][string]$SamHint,
    [Parameter(Mandatory)][System.Security.SecureString]$SecurePassword,
    [Parameter(Mandatory)][string]$UpnSuffix,
    [hashtable]$DepartmentGroupMap,
    [string]$ItDepartment,
    [string]$ItGroupName,
    [hashtable]$ServerParams
  )

  $result = [pscustomobject]@{
    SamAccountName = $null
    Created        = $false
    TargetOu       = $DomainDN
    GroupMessages  = New-Object System.Collections.Generic.List[string]
  }

  $nameTrimmed = $DisplayName.Trim()
  if ([string]::IsNullOrWhiteSpace($nameTrimmed)) {
    Write-Warning 'Alternate admin display name is blank; skipping creation.'
    return $result
  }

  $samBase = $SamHint
  if ([string]::IsNullOrWhiteSpace($samBase)) {
    $samBase = ($nameTrimmed -replace '\s+', '')
  }
  $samBase = ($samBase -replace '[^A-Za-z0-9]','').ToLower()
  if ([string]::IsNullOrWhiteSpace($samBase)) {
    $samBase = 'labadmin'
  }

  $serverParamsLocal = @{}
  if ($ServerParams) {
    foreach ($key in $ServerParams.Keys) {
      $serverParamsLocal[$key] = $ServerParams[$key]
    }
  }

  $userLookupParams = @{ LDAPFilter = "(sAMAccountName=$samBase)"; SearchBase = $DomainDN; ErrorAction = 'SilentlyContinue'; Properties = 'SamAccountName','DistinguishedName','SID' }
  $userLookupParams += $serverParamsLocal
  $existingUser = Get-ADUser @userLookupParams

  $finalSam = $samBase
  if (-not $existingUser) {
    $finalSam = New-UniqueSam $samBase
  } else {
    $finalSam = [string]$existingUser.SamAccountName
  }

  $result.SamAccountName = $finalSam

  $givenName = $nameTrimmed
  $surname   = $nameTrimmed
  $nameParts = $nameTrimmed -split '\s+'
  if ($nameParts.Count -gt 0) { $givenName = $nameParts[0] }
  if ($nameParts.Count -gt 1) { $surname = $nameParts[-1] }

  $targetOu = $DomainDN
  if ($ItDepartment) {
    $targetOu = "OU=$ItDepartment,$DomainDN"
  }
  $result.TargetOu = $targetOu

  $adminUser = $existingUser
  if (-not $adminUser) {
    $adminUpn = '{0}@{1}' -f $finalSam,$UpnSuffix
    $newAdminParams = @{ Name = $nameTrimmed; GivenName = $givenName; Surname = $surname; SamAccountName = $finalSam; UserPrincipalName = $adminUpn; Path = $targetOu; AccountPassword = $SecurePassword; Enabled = $true; ChangePasswordAtLogon = $false; ErrorAction = 'Stop' }
    $newAdminParams += $serverParamsLocal
    try {
      New-ADUser @newAdminParams | Out-Null
      Write-Host ("Created alternate admin account: {0}" -f $finalSam) -ForegroundColor Green
      $result.Created = $true
      $getAdminParams = @{ Identity = $finalSam; Properties = 'SamAccountName','DistinguishedName','SID' }
      $getAdminParams += $serverParamsLocal
      $adminUser = Get-ADUser @getAdminParams
    } catch {
      Write-Warning ("Failed to create alternate admin '{0}': {1}" -f $nameTrimmed,$_.Exception.Message)
      return $result
    }
  } else {
    Write-Host ("Alternate admin already exists: {0}" -f $finalSam) -ForegroundColor Cyan
  }

  if (-not $adminUser) {
    Write-Warning ("Unable to locate alternate admin account '{0}' after creation attempt." -f $finalSam)
    return $result
  }

  $enableParams = @{ Identity = $adminUser; ErrorAction = 'SilentlyContinue' }
  $enableParams += $serverParamsLocal
  try { Enable-ADAccount @enableParams } catch {}

  $changeParams = @{ Identity = $adminUser; ChangePasswordAtLogon = $false; ErrorAction = 'SilentlyContinue' }
  $changeParams += $serverParamsLocal
  try { Set-ADUser @changeParams } catch {}

  $groupTargets = New-Object System.Collections.Generic.List[pscustomobject]
  if ($DepartmentGroupMap -and $ItDepartment -and $DepartmentGroupMap.ContainsKey($ItDepartment)) {
    $itGroupInfo = $DepartmentGroupMap[$ItDepartment]
    if ($itGroupInfo -and $itGroupInfo.DistinguishedName) {
      $groupTargets.Add([pscustomobject]@{ Identity = $itGroupInfo.DistinguishedName; Label = $itGroupInfo.Name }) | Out-Null
    }
  } elseif ($ItGroupName) {
    $groupTargets.Add([pscustomobject]@{ Identity = $ItGroupName; Label = $ItGroupName }) | Out-Null
  }

  $groupTargets.Add([pscustomobject]@{ Identity = 'Domain Admins'; Label = 'Domain Admins' }) | Out-Null
  $groupTargets.Add([pscustomobject]@{ Identity = 'Administrators'; Label = 'Administrators' }) | Out-Null

  $seenGroups = New-Object System.Collections.Generic.HashSet[string]([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($groupTarget in $groupTargets) {
    if (-not $groupTarget) { continue }
    $identityValue = [string]$groupTarget.Identity
    if ([string]::IsNullOrWhiteSpace($identityValue)) { continue }
    if (-not $seenGroups.Add($identityValue)) { continue }

    $addGroupParams = @{ Identity = $identityValue; Members = $adminUser; ErrorAction = 'Stop' }
    $addGroupParams += $serverParamsLocal
    try {
      Add-ADGroupMember @addGroupParams
      $result.GroupMessages.Add(("Added to {0}" -f $groupTarget.Label)) | Out-Null
    } catch {
      if ($_.Exception.Message -match 'already a member') {
        $result.GroupMessages.Add(("Already a member of {0}" -f $groupTarget.Label)) | Out-Null
      } else {
        Write-Warning ("Failed to add alternate admin to {0}: {1}" -f $groupTarget.Label,$_.Exception.Message)
      }
    }
  }

  return $result
}

# Capture special department names for ACL logic (IT, Executives, Management).
$itDepartmentName   = ($Departments | Where-Object { $_ -match '^(?i)IT$' } | Select-Object -First 1)
$execDepartmentName = ($Departments | Where-Object { $_ -match '^(?i)Executives$' } | Select-Object -First 1)
$mgmtDepartmentName = ($Departments | Where-Object { $_ -match '^(?i)Management$' } | Select-Object -First 1)

if ($itDepartmentName) {
  $itDepartmentName = [string]$itDepartmentName
  $itGroupName = Resolve-DeptGroupName -Department $itDepartmentName
} else {
  $itDepartmentName = $null
  $itGroupName = $null
}

if ($execDepartmentName) {
  $execDepartmentName = [string]$execDepartmentName
  $execGroupName = Resolve-DeptGroupName -Department $execDepartmentName
} else {
  $execDepartmentName = $null
  $execGroupName = $null
}

if ($mgmtDepartmentName) {
  $mgmtDepartmentName = [string]$mgmtDepartmentName
  $mgmtGroupName = Resolve-DeptGroupName -Department $mgmtDepartmentName
} else {
  $mgmtDepartmentName = $null
  $mgmtGroupName = $null
}

# Build the share and folder structure when not explicitly skipped.
if (-not $SkipFoldersAndShares) {
  Ensure-Directory -Path $BasePath
  $homePath = Join-Path -Path $BasePath -ChildPath 'Home'
  $profilesPath = Join-Path -Path $BasePath -ChildPath 'Profiles'
  Ensure-Directory -Path $homePath
  Ensure-Directory -Path $profilesPath
  Ensure-Share -Name 'Home' -Path $homePath
  Ensure-Share -Name 'Profiles' -Path $profilesPath

  foreach ($dept in $Departments) {
    $folder = Join-Path -Path $BasePath -ChildPath $dept
    Ensure-Directory -Path $folder
    Ensure-Share -Name $dept -Path $folder
  }
}

# Ensure every department has a matching OU and group structure.
foreach ($dept in $Departments) {
  Ensure-OU $dept
  $grpCN = Resolve-DeptGroupName -Department $dept
  Ensure-Group $grpCN $dept
}

# Resolve department groups once so later lookups do not repeat LDAP queries.
$deptGroupMap = [System.Collections.Hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
foreach ($dept in $Departments) {
  $grpCN = Resolve-DeptGroupName -Department $dept
  $deptGroupLookup = @{ LDAPFilter = "(cn=$grpCN)"; SearchBase = "OU=$dept,$DomainDN"; ErrorAction = 'SilentlyContinue' }
  $deptGroupLookup += $adServerParams
  $deptGroupObj = Get-ADGroup @deptGroupLookup -Properties SID,SamAccountName
  if ($deptGroupObj) {
    $deptGroupMap[$dept] = [pscustomobject]@{
      Name = $grpCN
      Sam  = if ($deptGroupObj.SamAccountName) { [string]$deptGroupObj.SamAccountName } else { $grpCN }
      Sid  = if ($deptGroupObj.SID) { $deptGroupObj.SID.Value } else { $null }
      DistinguishedName = $deptGroupObj.DistinguishedName
    }
  }
}

# Ensure the alternate administrative account exists and has elevated memberships.
$alternateAdminSummary = Ensure-AlternateAdminAccount -DisplayName $AltAdminDisplayName -SamHint $AltAdminSamAccountName -SecurePassword $AltAdminSecurePwd -UpnSuffix $UPNSuffix -DepartmentGroupMap $deptGroupMap -ItDepartment $itDepartmentName -ItGroupName $itGroupName -ServerParams $adServerParams

# Lookup special security groups to reuse their SIDs for ACL grants.
$itGroupObject = $null
if ($itGroupName) {
  $itLookup = @{ Identity = $itGroupName; ErrorAction = 'SilentlyContinue'; Properties = 'SID','SamAccountName' }
  $itLookup += $adServerParams
  $itGroupObject = Get-ADGroup @itLookup
}

$itAclCandidates = @()
if ($itGroupObject) {
  $itSamResolved = if ($itGroupObject.SamAccountName) { [string]$itGroupObject.SamAccountName } else { $itGroupName }
  $itSidResolved = if ($itGroupObject.SID) { $itGroupObject.SID.Value } else { $null }
  $itAclCandidates = Get-IcaclsPrincipalCandidates -SamAccountName $itSamResolved -SidValue $itSidResolved
}

$execGroupObject = $null
if ($execGroupName) {
  $execLookup = @{ Identity = $execGroupName; ErrorAction = 'SilentlyContinue'; Properties = 'SID','SamAccountName' }
  $execLookup += $adServerParams
  $execGroupObject = Get-ADGroup @execLookup
}

$execAclCandidates = @()
if ($execGroupObject) {
  $execSamResolved = if ($execGroupObject.SamAccountName) { [string]$execGroupObject.SamAccountName } else { $execGroupName }
  $execSidResolved = if ($execGroupObject.SID) { $execGroupObject.SID.Value } else { $null }
  $execAclCandidates = Get-IcaclsPrincipalCandidates -SamAccountName $execSamResolved -SidValue $execSidResolved
}

$mgmtGroupObject = $null
if ($mgmtGroupName) {
  $mgmtLookup = @{ Identity = $mgmtGroupName; ErrorAction = 'SilentlyContinue'; Properties = 'SID','SamAccountName' }
  $mgmtLookup += $adServerParams
  $mgmtGroupObject = Get-ADGroup @mgmtLookup
}

$mgmtAclCandidates = @()
if ($mgmtGroupObject) {
  $mgmtSamResolved = if ($mgmtGroupObject.SamAccountName) { [string]$mgmtGroupObject.SamAccountName } else { $mgmtGroupName }
  $mgmtSidResolved = if ($mgmtGroupObject.SID) { $mgmtGroupObject.SID.Value } else { $null }
  $mgmtAclCandidates = Get-IcaclsPrincipalCandidates -SamAccountName $mgmtSamResolved -SidValue $mgmtSidResolved
}

# Apply default NTFS permissions across departmental folders.
if (-not $SkipAclForDepartments) {
  Write-Host "`nApplying department NTFS ACLs..." -ForegroundColor Yellow
  foreach ($dept in $Departments) {
    $folder = Join-Path $BasePath $dept
    if (-not (Test-Path -LiteralPath $folder)) {
      continue
    }
    $grpCN = Resolve-DeptGroupName -Department $dept
    $deptGroup = $null
    if ($deptGroupMap.ContainsKey($dept)) {
      $deptGroup = $deptGroupMap[$dept]
    }

    $deptPrincipalCandidates = @()
    if ($deptGroup) {
      $deptPrincipalCandidates = Get-IcaclsPrincipalCandidates -SamAccountName $deptGroup.Sam -SidValue $deptGroup.Sid
    }

    Invoke-IcaclsCommand -Arguments @($folder,'/inheritance:r') | Out-Null
    Invoke-IcaclsCommand -Arguments @($folder,'/grant','NT AUTHORITY\SYSTEM:(OI)(CI)F') -Description "SYSTEM full" | Out-Null
    Invoke-IcaclsCommand -Arguments @($folder,'/grant','BUILTIN\Administrators:(OI)(CI)F') -Description "Administrators full" | Out-Null
    if ($itAclCandidates.Count -gt 0) {
      Grant-IcaclsPermission -Path $folder -PrincipalCandidates $itAclCandidates -AccessRule '(OI)(CI)F' -Description 'IT full' | Out-Null
    } elseif ($itGroupName) {
      Write-Host ("IT group missing, skipped NTFS grant for: {0}" -f $itGroupName) -ForegroundColor Yellow
    }
    if ($deptPrincipalCandidates.Count -gt 0) {
      Grant-IcaclsPermission -Path $folder -PrincipalCandidates $deptPrincipalCandidates -AccessRule '(OI)(CI)F' -Description ("{0} full" -f $grpCN) | Out-Null
    } else {
      Write-Host ("Department group missing, skipped NTFS grant for: {0}" -f $grpCN) -ForegroundColor Yellow
    }

    $isItFolder = $false
    if ($itDepartmentName) {
      if ([string]::Equals($dept,$itDepartmentName,[System.StringComparison]::OrdinalIgnoreCase)) {
        $isItFolder = $true
      }
    }

    if (-not $isItFolder) {
      if ($execAclCandidates.Count -gt 0) {
        Grant-IcaclsPermission -Path $folder -PrincipalCandidates $execAclCandidates -AccessRule '(OI)(CI)RX' -Description 'Executives read' | Out-Null
      } elseif ($execGroupName) {
        Write-Host ("Executives group missing, skipped NTFS grant for: {0}" -f $execGroupName) -ForegroundColor Yellow
      }
      if ($mgmtAclCandidates.Count -gt 0) {
        Grant-IcaclsPermission -Path $folder -PrincipalCandidates $mgmtAclCandidates -AccessRule '(OI)(CI)RX' -Description 'Management read' | Out-Null
      } elseif ($mgmtGroupName) {
        Write-Host ("Management group missing, skipped NTFS grant for: {0}" -f $mgmtGroupName) -ForegroundColor Yellow
      }
    }

    Invoke-IcaclsCommand -Arguments @($folder,'/remove','BUILTIN\Users') -Description 'Remove BUILTIN\Users' | Out-Null
    Write-Host ("  NTFS set: {0}" -f $folder) -ForegroundColor Green
  }
}

# Load CSV rows only when an import file is supplied.
$rows = @()
if ($CsvPath) {
  $rows = Import-Csv -LiteralPath $CsvPath
}

$createdUsers = 0
$updatedUsers = 0
$malformedLogPath = $null

# Import users when rows are present in the CSV.
if ($rows.Count -gt 0) {
  $bad = New-Object System.Collections.Generic.List[string]

  Write-Host "`nImporting users..." -ForegroundColor Yellow

  $headers  = $rows[0].psobject.Properties.Name
  $FullHdr  = ($headers | Where-Object { $_ -match '^(FullName|Name)$' } | Select-Object -First 1)
  $FirstHdr = ($headers | Where-Object { $_ -match '^(FirstName|First Name|GivenName|Given Name|FName)$' } | Select-Object -First 1)
  $LastHdr  = ($headers | Where-Object { $_ -match '^(LastName|Last Name|Surname|Sur Name|LName)$' } | Select-Object -First 1)

  if (-not $FullHdr -and (-not $FirstHdr -or -not $LastHdr)) {
    $firstHeader = $headers | Select-Object -First 1
    if ($firstHeader) {
      $FullHdr = [string]$firstHeader
      Write-Host ("Treating column '{0}' as FullName (fallback)." -f $FullHdr) -ForegroundColor Cyan
    } else {
      throw 'CSV must have FullName OR First/Last columns.'
    }
  }

  foreach ($row in $rows) {
    $First = $null
    $Last  = $null

    if ($FullHdr) {
      $fullValue = ([string]$row.$FullHdr).Trim()
      if (-not $fullValue) {
        continue
      }
      $parts = $fullValue -split '\s+'
      if ($parts.Count -lt 2) {
        $bad.Add($fullValue)
        continue
      }
      $First = $parts[0]
      $Last  = $parts[-1]
    } else {
      $First = [string]$row.$FirstHdr
      $Last  = [string]$row.$LastHdr
      if ([string]::IsNullOrWhiteSpace($First) -or [string]::IsNullOrWhiteSpace($Last)) {
        $bad.Add("$First $Last")
        continue
      }
      $First = $First.Trim()
      $Last  = $Last.Trim()
    }

    $baseSam = if ($First.Length -gt 0) { $First.Substring(0,1) + $Last } else { $Last }
    $sam = New-UniqueSam $baseSam
    if (-not $sam) {
      $bad.Add("$First $Last")
      continue
    }

    $upn = "$sam@$UPNSuffix"
    $dept = Get-Random $Departments
    $ouPath = "OU=$dept,$DomainDN"
    $grpCN = Resolve-DeptGroupName -Department $dept
    $groupDn = "CN=$grpCN,$ouPath"
    if ($deptGroupMap.ContainsKey($dept)) {
      $groupDn = $deptGroupMap[$dept].DistinguishedName
    }

    $userObj = $null
    try {
      $newUserParams = @{ Name = ("{0} {1}" -f $First,$Last); GivenName = $First; Surname = $Last; SamAccountName = $sam; UserPrincipalName = $upn; Path = $ouPath; AccountPassword = $SecurePwd; Enabled = $true; ChangePasswordAtLogon = $true; ErrorAction = 'Stop' }
      $newUserParams += $adServerParams
      New-ADUser @newUserParams

      $getUserParams = @{ Identity = $sam; Properties = 'SID','SamAccountName' }
      $getUserParams += $adServerParams
      $userObj = Get-ADUser @getUserParams
      $createdUsers++
    } catch {
      if ($_.Exception.Message -match 'already in use' -or $_.FullyQualifiedErrorId -match 'ActiveDirectoryServer:8305') {
        $fullName = "{0} {1}" -f $First,$Last
        $escapedFullName = $fullName.Replace("'","''")
        $existingUserParams = @{ Filter = "Name -eq '$escapedFullName'"; SearchBase = $DomainDN; ErrorAction = 'SilentlyContinue'; Properties = 'SID','SamAccountName' }
        $existingUserParams += $adServerParams
        $userObj = Get-ADUser @existingUserParams
        if ($userObj) {
          $updatedUsers++
        } else {
          Write-Warning ("Duplicate CN but user not found: {0} {1}" -f $First,$Last)
          continue
        }
      } else {
        Write-Warning ("Failed to create '{0} {1}': {2}" -f $First,$Last,$_.Exception.Message)
        continue
      }
    }

    if ($userObj) {
      $changePwdParams = @{ Identity = $userObj; ChangePasswordAtLogon = $true; ErrorAction = 'SilentlyContinue' }
      $changePwdParams += $adServerParams
      try { Set-ADUser @changePwdParams } catch {}

      $addGroupParams = @{ Identity = $groupDn; Members = $userObj; ErrorAction = 'SilentlyContinue' }
      $addGroupParams += $adServerParams
      try { Add-ADGroupMember @addGroupParams } catch {}

      $homeUNC   = "\\\\$ServerName\\Home\\$sam"
      $homeLocal = Join-Path (Join-Path $BasePath 'Home') $sam

      $homeParams = @{ Identity = $userObj; HomeDrive = 'H:'; HomeDirectory = $homeUNC; ErrorAction = 'SilentlyContinue' }
      $homeParams += $adServerParams
      try { Set-ADUser @homeParams } catch {}

      $homeFolderCreated = $false
      if (-not (Test-Path -LiteralPath $homeLocal)) {
        New-Item -ItemType Directory -Path $homeLocal -Force | Out-Null
        $homeFolderCreated = $true
      }

      if ($homeFolderCreated) {
        Write-Host ("Ensured home folder: {0} (created)" -f $homeLocal) -ForegroundColor Green
      } else {
        Write-Host ("Ensured home folder: {0} (already existed)" -f $homeLocal) -ForegroundColor Cyan
      }

      if (Test-Path -LiteralPath $homeLocal) {
        $userSidValue = $null
        if ($userObj.SID) { $userSidValue = $userObj.SID.Value }
        $userPrincipalCandidates = Get-IcaclsPrincipalCandidates -SamAccountName $sam -SidValue $userSidValue

        Invoke-IcaclsCommand -Arguments @($homeLocal,'/inheritance:r') | Out-Null
        if ($userPrincipalCandidates.Count -gt 0) {
          Grant-IcaclsPermission -Path $homeLocal -PrincipalCandidates $userPrincipalCandidates -AccessRule '(OI)(CI)F' -Description ("Home full for {0}" -f $sam) | Out-Null
        } else {
          Write-Host ("Skipping user home ACL because SID/SAM missing for: {0}" -f $sam) -ForegroundColor Yellow
        }

        if ($itAclCandidates.Count -gt 0) {
          Grant-IcaclsPermission -Path $homeLocal -PrincipalCandidates $itAclCandidates -AccessRule '(OI)(CI)F' -Description 'IT home full' | Out-Null
        } elseif ($itGroupName) {
          Write-Host ("IT group missing, skipped home-folder grant for: {0}" -f $itGroupName) -ForegroundColor Yellow
        }

        Invoke-IcaclsCommand -Arguments @($homeLocal,'/grant','NT AUTHORITY\SYSTEM:(OI)(CI)F') -Description 'SYSTEM home full' | Out-Null
      }
    }
  }

  # Persist any malformed names to the desktop for follow-up.
  if ($bad.Count -gt 0) {
    $malformedLogPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'MalformedNames.txt'
    $bad | Set-Content -LiteralPath $malformedLogPath -Encoding UTF8
    Write-Host ("Some rows were skipped. See: {0}" -f $malformedLogPath) -ForegroundColor Yellow
  }

  Write-Host ("Created {0} users; updated {1} existing users." -f $createdUsers,$updatedUsers) -ForegroundColor Green
} elseif ($CsvPath) {
  Write-Warning 'The provided file did not contain any rows to import.'
}

# Emit a verification summary for shares, OUs, groups, and user import status.
Write-Host "`nVERIFY:" -ForegroundColor Cyan
Write-Host ("  Shares on {0}:" -f $ServerName) -ForegroundColor Cyan
(Get-SmbShare | Where-Object { $_.Name -in ($Departments + 'Home' + 'Profiles') }) |
  Select-Object Name,Path | Format-Table -AutoSize

Write-Host "`n  OUs:" -ForegroundColor Cyan
$ouFilterParts = $Departments | ForEach-Object { "(ou=$_)" }
$ouFilter = "(|" + ($ouFilterParts -join '') + ")"
$verifyOuParams = @{ SearchBase = $DomainDN; LDAPFilter = $ouFilter }
$verifyOuParams += $adServerParams
Get-ADOrganizationalUnit @verifyOuParams |
  Select-Object Name | Sort-Object Name | Format-Table -AutoSize

Write-Host "`n  Groups (by OU):" -ForegroundColor Cyan
foreach ($dept in $Departments) {
  $grpCN = Resolve-DeptGroupName -Department $dept
  $verifyGroupParams = @{ LDAPFilter = "(cn=$grpCN)"; SearchBase = "OU=$dept,$DomainDN"; ErrorAction = 'SilentlyContinue' }
  $verifyGroupParams += $adServerParams
  $grp = Get-ADGroup @verifyGroupParams
  if ($grp) {
    Write-Host ("{0,-18} -> {1}" -f $dept, $grp.Name)
  } else {
    Write-Host ("{0,-18} -> <missing>" -f $dept)
  }
}

Write-Host "`n  Alternate admin account:" -ForegroundColor Cyan
if ($alternateAdminSummary -and $alternateAdminSummary.SamAccountName) {
  $accountColor = if ($alternateAdminSummary.Created) { 'Green' } else { 'Cyan' }
  Write-Host ("    sAMAccountName: {0}" -f $alternateAdminSummary.SamAccountName) -ForegroundColor $accountColor
  Write-Host ("    Target OU: {0}" -f $alternateAdminSummary.TargetOu) -ForegroundColor $accountColor
  if ($alternateAdminSummary.GroupMessages -and $alternateAdminSummary.GroupMessages.Count -gt 0) {
    foreach ($msg in $alternateAdminSummary.GroupMessages) {
      Write-Host ("    {0}" -f $msg)
    }
  } else {
    Write-Host '    No group membership updates were applied.' -ForegroundColor Yellow
  }
} else {
  Write-Host '    Alternate admin account information unavailable.' -ForegroundColor Yellow
}

Write-Host "`n  User import summary:" -ForegroundColor Cyan
if ($CsvPath) {
  Write-Host ("    Created: {0}" -f $createdUsers)
  Write-Host ("    Updated: {0}" -f $updatedUsers)
  if ($malformedLogPath) {
    Write-Host ("    Malformed entries logged to: {0}" -f $malformedLogPath) -ForegroundColor Yellow
  } else {
    Write-Host '    No malformed entries logged.'
  }
} else {
  Write-Host '    Import skipped (no CSV supplied).'
}

Write-Host "`nDone." -ForegroundColor Green
