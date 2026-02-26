<#!
.SYNOPSIS
    Autopilot Assistant GUI for HWID collection and Intune Autopilot import.

.DESCRIPTION
    This script provides a WPF-based workflow to:
    - Collect local device hardware hash (HWID) to CSV
    - Connect to Microsoft Graph (interactive delegated sign-in)
    - Upload/import devices into Windows Autopilot
    - Retry failed rows from the latest upload attempt
    - Show clear upload outcomes and diagnostics

    The tool is intended for IT operations. Review required permissions,
    execution context, and network/proxy prerequisites before production use.

    Exit codes:
    - Exit 0: Completed successfully
    - Exit 1: Failed or requires further action

.RUN AS
    User or System (based on assignment model and script dependencies).

.EXAMPLE
    .\Autopilot Assistant.ps1

.NOTES
    Author  : Mohammad Abdelkader
    Website : momar.tech
    Date    : 2026-02-25
    Version : 2.0
#>

#region ============================== BOOTSTRAPPING / .NET ==========================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase | Out-Null
Add-Type -AssemblyName System.Windows.Forms | Out-Null
#endregion ============================== BOOTSTRAPPING / .NET ========================================

#region ============================== CONSTANTS / BRANDING / APP PATHS ==============================
$AppVersion  = '2.0'


# Delegated scopes for interactive Graph auth (sufficient for Autopilot R/W)
$DefaultScopes = @(
  'User.Read',
  'Device.Read.All',
  'DeviceManagementServiceConfig.ReadWrite.All'
)

# Root working directory on system drive
$BasePath = Join-Path $env:SystemDrive 'AutopilotAssistant'
$Paths = @{
  Root     = $BasePath
  HwId     = Join-Path $BasePath 'HWID'
  Logs     = Join-Path $BasePath 'Logs'
}
$Paths.GetEnumerator() | ForEach-Object {
  if (-not (Test-Path -LiteralPath $_.Value)) {
    New-Item -ItemType Directory -Path $_.Value -Force | Out-Null
  }
}

# Per-day log file
$LogFile = Join-Path $Paths.Logs ("app_{0}.log" -f (Get-Date -Format 'yyyyMMdd'))
#endregion ============================== CONSTANTS / BRANDING / APP PATHS ============================

#region ============================== UTILITY HELPERS (UI + LOGGING) ================================
function New-Brush([string]$color) {
  $fallback = [Windows.Media.Brushes]::Black
  if ([string]::IsNullOrWhiteSpace($color)) { return $fallback }

  $text = $color.Trim()

  # Normalize short hex notations (#RGB / #ARGB) to full forms to reduce converter errors.
  if ($text -match '^#([0-9A-Fa-f]{3})$') {
    $h = $Matches[1]
    $text = ('#{0}{0}{1}{1}{2}{2}' -f $h[0], $h[1], $h[2])
  } elseif ($text -match '^#([0-9A-Fa-f]{4})$') {
    $h = $Matches[1]
    $text = ('#{0}{0}{1}{1}{2}{2}{3}{3}' -f $h[0], $h[1], $h[2], $h[3])
  }

  try {
    $bc = New-Object Windows.Media.BrushConverter
    $b = $bc.ConvertFromInvariantString($text)
    if ($b) { return $b }
  } catch { }

  try {
    $c = [Windows.Media.ColorConverter]::ConvertFromString($text)
    if ($c -is [Windows.Media.Color]) {
      $sb = New-Object Windows.Media.SolidColorBrush($c)
      try { $sb.Freeze() } catch { }
      return $sb
    }
  } catch { }

  return $fallback
}

# Function: Add-Log
# Writes color-coded entries to Message Center and file log.
function Add-Log([string]$text, [string]$level = 'INFO') {

  # Foreground colors tuned for dark console background (#1F2D3A)
  $lvl = if ([string]::IsNullOrWhiteSpace($level)) { 'INFO' } else { $level.ToUpperInvariant() }
  $fg = '#E5E7EB'  # default (gray-200)
  switch ($lvl) {
    'SUCCESS' { $fg = '#34D399' } # emerald
    'WARN'    { $fg = '#FBBF24' } # amber
    'ERROR'   { $fg = '#F87171' } # red
    default   { $fg = '#93C5FD' } # blue
  }

  if ($Window -and $TxtLog) {
    $Window.Dispatcher.Invoke([action]{
      $p = New-Object Windows.Documents.Paragraph
      $p.Margin = '0,0,0,4'

      $run = New-Object Windows.Documents.Run ("[{0}] {1}" -f $lvl, $text)
      $run.Foreground = New-Brush $fg
      [void]$p.Inlines.Add($run)

      $TxtLog.Document.Blocks.Add($p)
      $TxtLog.ScrollToEnd()
    })
  }

  try {
    Add-Content -LiteralPath $LogFile -Value ("[{0}] {1}" -f $lvl, $text)
  } catch { }
}

function Show-UploadBlockedDeviceExists {
  param(
    [string[]]$Serials,
    [string]$WindowTitleText = 'Device Already Exists',
    [string]$HeaderTitle = 'Device Already Registered',
    [string]$HeaderSubtitle = 'This device already exists in Windows Autopilot.',
    [string]$InstructionText = 'No upload is needed. Delete the device first only if you want to re-import it.',
    [string]$FallbackBody = 'Device already exists in Autopilot and was skipped.'
  )

  $lines = @()
  foreach ($s in @($Serials)) {
    if (-not [string]::IsNullOrWhiteSpace($s)) { $lines += ("- " + $s) }
  }
  if ($lines.Count -eq 0) { $lines = @('- Unknown serial') }
  $x = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$WindowTitleText"
        Width="560" Height="360"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#F6F8FB"
        FontFamily="Segoe UI"
        FontSize="13">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="12"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="#FEF3C7" BorderBrush="#FCD34D" BorderThickness="1" CornerRadius="6" Padding="10">
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="!" FontSize="18" FontWeight="Bold" Margin="0,0,10,0"/>
          <StackPanel>
          <TextBlock Text="$HeaderTitle" FontSize="18" FontWeight="Bold" Foreground="#92400E"/>
          <TextBlock Text="$HeaderSubtitle"
                     Margin="0,4,0,0" Foreground="#7C2D12"/>
          </StackPanel>
        </StackPanel>
    </Border>

    <Border Grid.Row="2" Background="White" BorderBrush="#E4E9F0" BorderThickness="1" CornerRadius="6" Padding="12">
      <Grid>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="8"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0"
                   Text="$InstructionText"
                   Foreground="#334155" FontWeight="SemiBold"/>
        <Border Grid.Row="2" Background="#F8FAFD" BorderBrush="#E2E8F0" BorderThickness="1" CornerRadius="5" Padding="10">
          <DockPanel>
            <TextBlock DockPanel.Dock="Top" Text="Serial Number:" FontWeight="SemiBold" Foreground="#0F172A" Margin="0,0,0,8"/>
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <TextBlock x:Name="TxtSerials" FontFamily="Consolas" Foreground="#1E293B"/>
            </ScrollViewer>
          </DockPanel>
        </Border>
      </Grid>
    </Border>

    <Grid Grid.Row="4">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="130"/>
      </Grid.ColumnDefinitions>
      <Button Grid.Column="1" x:Name="BtnOk" Content="OK"
              Height="34" Background="#D8E5FF" Foreground="#1E3A6D"
              BorderBrush="#B7CBEF" BorderThickness="1"/>
    </Grid>
  </Grid>
</Window>
"@
  try {
    $w = [Windows.Markup.XamlReader]::Parse($x)
    $TxtSerials = $w.FindName('TxtSerials')
    $BtnOk = $w.FindName('BtnOk')
    $TxtSerials.Text = ($lines -join [Environment]::NewLine)
    $BtnOk.Add_Click({ $w.Close() })
    [void]$w.ShowDialog()
  } catch {
    # fallback
    try {
      [void][System.Windows.MessageBox]::Show(
        ("Upload note.`r`n`r`n" + $FallbackBody + "`r`n`r`n" + ($lines -join "`r`n")),
        $WindowTitleText,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
      )
    } catch { }
  }
}

function Start-UploadProgress {
  if ($ProgressBar) {
    $script:UploadProgressActive = $true
    $ProgressBar.Value = 8
    $ProgressBar.Visibility = 'Visible'
  } else {
    $script:UploadProgressActive = $true
  }
}

function Step-UploadProgress {
  if ($ProgressBar -and $script:UploadProgressActive) {
    $next = [double]$ProgressBar.Value + 1.5
    if ($next -gt 92) { $next = 12 }
    $ProgressBar.Value = $next
  }
}

function Stop-UploadProgress {
  $script:UploadProgressActive = $false
  if ($ProgressBar) {
    $ProgressBar.Value = 0
    $ProgressBar.Visibility = 'Collapsed'
  }
}

function Ensure-RequiredModule {
  param(
    [Parameter(Mandatory)][string]$ModuleName,
    [Parameter(Mandatory)][string]$ReadyFlagName,
    [switch]$InstallIfMissing
  )

  try {
    $readyVar = Get-Variable -Scope Script -Name $ReadyFlagName -ErrorAction SilentlyContinue
    if ($readyVar -and [bool]$readyVar.Value -and (Get-Module -Name $ModuleName)) {
      return $true
    }

    if (Get-Module -ListAvailable -Name $ModuleName) {
      Import-Module $ModuleName -ErrorAction Stop
      Set-Variable -Scope Script -Name $ReadyFlagName -Value $true
      return $true
    }

    if (-not $InstallIfMissing) {
      Add-Log "$ModuleName module is not installed." "WARN"
      return $false
    }

    if ($env:USERNAME -eq 'SYSTEM') {
      Add-Log "$ModuleName is missing and cannot be installed in SYSTEM context. Install it beforehand." "ERROR"
      return $false
    }

    Add-Log "$ModuleName not found. Installing for current user..." "WARN"

    if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
      Add-Log "Install-Module command is not available (PowerShellGet missing)." "ERROR"
      return $false
    }

    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
      Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
    }

    try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue | Out-Null } catch { }

    Install-Module -Name $ModuleName -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
    Import-Module $ModuleName -ErrorAction Stop
    Set-Variable -Scope Script -Name $ReadyFlagName -Value $true
    Add-Log "$ModuleName installed successfully." "SUCCESS"
    return $true
  } catch {
    Add-Log ("$ModuleName install/check failed: " + $_.Exception.Message) "ERROR"
    return $false
  }
}

# Function: Ensure-MgGraphModule
function Ensure-MgGraphModule {
  param([switch]$InstallIfMissing)
  return (Ensure-RequiredModule -ModuleName 'Microsoft.Graph.Authentication' -ReadyFlagName 'MgGraphModuleReady' -InstallIfMissing:$InstallIfMissing)
}

# Function: Ensure-MsalModule
function Ensure-MsalModule {
  param([switch]$InstallIfMissing)
  return (Ensure-RequiredModule -ModuleName 'MSAL.PS' -ReadyFlagName 'MsalModuleReady' -InstallIfMissing:$InstallIfMissing)
}

# Function: Sync-GraphContext
# Reads current Graph context and synchronizes UI state.
function Sync-GraphContext {
  param([switch]$Silent)

  $ctx = $null
  try { $ctx = Get-MgContext } catch { $ctx = $null }

  $script:GraphConnected = ($ctx -and ($ctx.Account -or $ctx.ClientId))
  $script:GraphAccount   = $null
  $script:GraphTenantId  = $null
  if ($ctx -and $ctx.Account)  { $script:GraphAccount = $ctx.Account }
  if ($ctx -and $ctx.TenantId) { $script:GraphTenantId = $ctx.TenantId }

  if (-not $Silent) {
    if ($script:GraphConnected) {
      $who = if ($script:GraphAccount) { $script:GraphAccount } else { 'app-only connection' }
      Add-Log ("Microsoft Graph is connected: " + $who) "SUCCESS"
    } else {
      Add-Log "Microsoft Graph is not connected. Use Connect first." "WARN"
    }
  }

  try { Update-MainGraphUI } catch { }
  return $script:GraphConnected
}

# Function: Connect-GraphBrowserInteractive
function Connect-GraphBrowserInteractive {
  try {
    if ($script:Async -and -not $script:Async.IsCompleted) {
      Add-Log "Another operation is running. Wait until it finishes." "WARN"
      return
    }

    # Fast path: skip auth flow if context is already connected.
    if (Sync-GraphContext -Silent) {
      Add-Log "Graph is already connected." "INFO"
      return
    }

    Add-Log "Starting Microsoft Graph browser interactive sign-in..." "INFO"

    if (-not (Ensure-MgGraphModule -InstallIfMissing)) {
      throw 'Microsoft.Graph.Authentication is not available.'
    }
    if (-not (Ensure-MsalModule -InstallIfMissing)) {
      throw 'MSAL.PS is not available.'
    }

    # Reconnect fast path: reuse cached access token from this app session.
    if (-not [string]::IsNullOrWhiteSpace($script:CachedGraphAccessToken)) {
      Add-Log "Trying cached Graph session token..." "INFO"
      try {
        $cachedSecureToken = ConvertTo-SecureString -String ([string]$script:CachedGraphAccessToken) -AsPlainText -Force
        Connect-MgGraph -AccessToken $cachedSecureToken -NoWelcome -ErrorAction Stop | Out-Null
        if (Sync-GraphContext -Silent) {
          Add-Log "Connected to Graph using cached session token." "SUCCESS"
          return
        }
      } catch {
        Add-Log "Cached token reuse failed. Falling back to browser sign-in..." "WARN"
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch { }
        $script:CachedGraphAccessToken = $null
      }
    }

    $script:GraphConnecting = $true
    Update-MainGraphUI
    $script:CurrentAction = 'graphauth'
    Add-Log "Trying cached Graph session first; browser opens only if needed..." "INFO"

    $script:PS = [System.Management.Automation.PowerShell]::Create()
    $script:PS.RunspacePool = $pool
    $script:PS.AddScript($GraphTokenWorker).AddArgument($DefaultScopes) | Out-Null
    $script:Async = $script:PS.BeginInvoke()
    $Timer.Start()
    Add-Log "Browser sign-in flow started in background..." "INFO"
  } catch {
    $script:GraphConnecting = $false
    if ($BtnGraphConnect) { $BtnGraphConnect.IsEnabled = $true }
    if ($BtnGraphDisconnect) { $BtnGraphDisconnect.IsEnabled = [bool]$script:GraphConnected }
    Add-Log ("Graph connect error: " + $_.Exception.Message) "ERROR"
    Update-MainGraphUI
  }
}

# Function: Disconnect-GraphDirect
# Disconnects active Graph context but keeps session token cache in memory
# for fast reconnect during the same app session.
function Disconnect-GraphDirect {
  try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 150
    $script:GraphConnecting = $false
    $script:GraphConnected = $false
    $script:GraphAccount = $null
    $script:GraphTenantId = $null
    Update-MainGraphUI
    Add-Log "Disconnected from Graph." "INFO"
  } catch {
    Add-Log ("Graph disconnect error: " + $_.Exception.Message) "ERROR"
  }
}

#endregion ============================== UTILITY HELPERS (UI + LOGGING) ==============================

#region ============================== DEVICE / ENVIRONMENT CHECKS ===================================
function Test-Internet {
  try {
    $tcp = New-Object System.Net.Sockets.TcpClient
    $ar  = $tcp.BeginConnect('www.microsoft.com', 443, $null, $null)
    if (-not $ar.AsyncWaitHandle.WaitOne(2000, $false)) { $tcp.Close(); return $false }
    $tcp.EndConnect($ar); $tcp.Close(); return $true
  } catch { return $false }
}

# Function: Refresh-DeviceInfo
# Refreshes local hardware and readiness indicators shown in UI.
function Refresh-DeviceInfo {
  $LblDevModel.Text = '-'
  $LblDevName.Text  = '-'
  $LblManufacturer.Text = '-'
  $LblSerial.Text   = '-'
  $LblFreeGb.Text   = '-'
  $LblTpm.Text      = 'Unknown'; $LblTpm.Foreground = New-Brush '#444'
  $LblNet.Text      = 'Not connected'; $LblNet.Foreground = New-Brush '#D13438'

  try {
    $cs   = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
    $bios = Get-CimInstance Win32_BIOS          -ErrorAction SilentlyContinue
    if ($cs -and $cs.Model)        { $LblDevModel.Text = $cs.Model }
    if ($cs -and $cs.Name)         { $LblDevName.Text  = $cs.Name }
    if ($cs -and $cs.Manufacturer) { $LblManufacturer.Text = $cs.Manufacturer }
    if ($bios -and $bios.SerialNumber) { $LblSerial.Text = $bios.SerialNumber }
  } catch { }

  try {
    $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction SilentlyContinue
    $free = 0
    if ($disks) { foreach ($d in $disks) { if ($d.FreeSpace) { $free += [int64]$d.FreeSpace } } }
    $LblFreeGb.Text = [string]([math]::Round($free/1GB, 0))
  } catch { }

  try {
    $tpm = Get-CimInstance -Namespace root\cimv2\security\microsofttpm -Class Win32_Tpm -ErrorAction Stop
    $spec = $null
    if ($tpm -and $tpm.SpecVersion)          { $spec = [string]$tpm.SpecVersion }
    elseif ($tpm -and $tpm.ManufacturerVersion) { $spec = [string]$tpm.ManufacturerVersion }
    if ($spec) {
      $LblTpm.Text = $spec
      if ($spec -match '2\.0') { $LblTpm.Foreground = New-Brush '#0A8A0A' }
      elseif ($spec -match '1\.2') { $LblTpm.Foreground = New-Brush '#D13438' }
      else { $LblTpm.Foreground = New-Brush '#444' }
    } else {
      $LblTpm.Text = 'Not present'
      $LblTpm.Foreground = New-Brush '#D13438'
    }
  } catch {
    $LblTpm.Text = 'Unknown'
    $LblTpm.Foreground = New-Brush '#444'
  }

  if (Test-Internet) { $LblNet.Text = 'Connected'; $LblNet.Foreground = New-Brush '#0A8A0A' }

}

# Function: Refresh-SessionInfo
function Refresh-SessionInfo {
  if ($LblSessionMachine) { $LblSessionMachine.Text = $env:COMPUTERNAME }

  $sessionUser = '-'
  try {
    $sessionUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  } catch { }
  if ($LblSessionUser) { $LblSessionUser.Text = $sessionUser }

  $elevationText = 'Unknown'
  try {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object System.Security.Principal.WindowsPrincipal($id)
    if ($pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
      $elevationText = 'Administrator'
    } else {
      $elevationText = 'Standard User'
    }
  } catch { }

  if ($LblSessionElevation) {
    $LblSessionElevation.Text = $elevationText
    if ($elevationText -eq 'Administrator') {
      $LblSessionElevation.Foreground = New-Brush '#166534'
    } elseif ($elevationText -eq 'Standard User') {
      $LblSessionElevation.Foreground = New-Brush '#92400E'
    } else {
      $LblSessionElevation.Foreground = New-Brush '#4B5563'
    }
  }
}

# Function: Test-IsAdmin
function Test-IsAdmin {
  try {
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object System.Security.Principal.WindowsPrincipal($id)
    return $pr.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch {
    return $false
  }
}

# Function: Run-PreChecks
# Validates required prerequisites before collect/upload operations.
function Run-PreChecks {
  param([ValidateSet('collect','enroll')][string]$Mode)

  $ok = $true

  if (-not (Test-IsAdmin)) {
    Add-Log "Pre-check failed: run the tool as Administrator." "ERROR"
    $ok = $false
  }

  try { Get-CimInstance Win32_BIOS -ErrorAction Stop | Out-Null }
  catch {
    Add-Log "Pre-check failed: Win32_BIOS is not reachable." "ERROR"
    $ok = $false
  }

  try { Get-CimInstance -Namespace root\cimv2\mdm\dmmap -Class MDM_DevDetail_Ext01 -ErrorAction Stop | Out-Null }
  catch {
    Add-Log "Pre-check failed: Autopilot hardware hash provider is not available." "ERROR"
    $ok = $false
  }

  if ($Mode -eq 'enroll') {
    if (-not (Test-Internet)) {
      Add-Log "Pre-check failed: internet connection is required for upload." "ERROR"
      $ok = $false
    }

    Add-Log "Pre-check: checking Microsoft Graph module..." "INFO"
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
      Add-Log "Microsoft.Graph.Authentication not found. Downloading from PSGallery..." "WARN"
    }
    if (-not (Ensure-MgGraphModule -InstallIfMissing)) { $ok = $false }

    Add-Log "Pre-check: checking Microsoft Graph connection..." "INFO"
    if (-not (Sync-GraphContext -Silent)) {
      Add-Log "Pre-check failed: connect to Microsoft Graph first." "ERROR"
      $ok = $false
    }
  }

  if ($ok) {
    Add-Log ("Pre-check passed for " + $Mode + ".") "SUCCESS"
  }
  return $ok
}

function Get-SerialFromRecord {
  param($Record)
  if (-not $Record) { return $null }
  $s = $null
  if ($Record.PSObject.Properties.Match('serialNumber').Count -gt 0) { $s = [string]$Record.serialNumber }
  elseif ($Record.PSObject.Properties.Match('Device Serial Number').Count -gt 0) { $s = [string]$Record.'Device Serial Number' }
  if ([string]::IsNullOrWhiteSpace($s)) { return $null }
  return $s.Trim()
}

function Get-UploadSerialCandidates {
  param(
    [string]$Path,
    [string]$FallbackSerial
  )

  $serials = @()
  try {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path) -or (Test-Path -LiteralPath $Path -PathType Container)) {
      if (-not [string]::IsNullOrWhiteSpace($FallbackSerial) -and $FallbackSerial -ne '-') {
        $serials += $FallbackSerial.Trim()
      }
    } elseif ($Path.ToLower().EndsWith('.json')) {
      $json = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
      foreach ($r in @($json)) {
        $s = Get-SerialFromRecord -Record $r
        if ($s) { $serials += $s }
      }
    } else {
      $csv = Import-Csv -LiteralPath $Path -ErrorAction Stop
      foreach ($r in @($csv)) {
        $s = Get-SerialFromRecord -Record $r
        if ($s) { $serials += $s }
      }
    }
  } catch {
    Add-Log ("Pre-check serial parse warning: " + $_.Exception.Message) "WARN"
  }

  $seen = @{}
  $out  = @()
  foreach ($s in @($serials)) {
    $k = $s.ToUpperInvariant()
    if (-not $seen.ContainsKey($k)) {
      $seen[$k] = $true
      $out += $s
    }
  }
  return @($out)
}

function Test-AutopilotSerialExists {
  param([Parameter(Mandatory)][string]$Serial)

  function Find-SerialInAutopilotPages {
    param(
      [Parameter(Mandatory)][string]$TargetSerial,
      [int]$MaxPages = 25
    )

    $next = "/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$top=100"
    $page = 0

    while ($next -and $page -lt $MaxPages) {
      $res = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop

      foreach ($item in @($res.value)) {
        $sn = [string]$item.serialNumber
        if (-not [string]::IsNullOrWhiteSpace($sn) -and $sn.Trim().ToUpperInvariant() -eq $TargetSerial) {
          return $true
        }
      }

      $next = $null
      if ($res.PSObject.Properties.Match('@odata.nextLink').Count -gt 0 -and $res.'@odata.nextLink') {
        $next = [string]$res.'@odata.nextLink'
      }
      $page++
    }

    return $false
  }

  if ([string]::IsNullOrWhiteSpace($Serial)) {
    return [pscustomobject]@{
      Exists = $false
      Blocking = $false
      Source = ''
      Serial = ''
    }
  }

  $targetSerial = $Serial.Trim().ToUpperInvariant()
  $safeSerial = $Serial.Replace("'", "''")
  $filter = "serialNumber eq '$safeSerial'"
  $f = [System.Uri]::EscapeDataString($filter)

  try {
    $uri1 = "/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=$f&`$top=1"
    $r1 = Invoke-MgGraphRequest -Method GET -Uri $uri1 -ErrorAction Stop
    if ($r1.value -and $r1.value.Count -gt 0) {
      return [pscustomobject]@{
        Exists = $true
        Blocking = $true
        Source = 'windowsAutopilotDeviceIdentities'
        Serial = $Serial
      }
    }
  } catch {
    $msg = [string]$_.Exception.Message
    Add-Log ("Pre-check Graph lookup warning for serial " + $Serial + " (windowsAutopilotDeviceIdentities): " + $msg) "WARN"
    try {
      Add-Log ("Pre-check: fallback serial scan started for " + $Serial + "...") "INFO"
      if (Find-SerialInAutopilotPages -TargetSerial $targetSerial) {
        return [pscustomobject]@{
          Exists = $true
          Blocking = $true
          Source = 'windowsAutopilotDeviceIdentities'
          Serial = $Serial
        }
      }
    } catch {
      Add-Log ("Pre-check Graph fallback scan warning for serial " + $Serial + ": " + $_.Exception.Message) "WARN"
    }
  }

  return [pscustomobject]@{
    Exists = $false
    Blocking = $false
    Source = ''
    Serial = $Serial
  }
}

# Function: Show-EnrollResultsWindow
# Shows a compact upload outcome window (success/failed/blocked/queued).
function Show-EnrollResultsWindow {
  param(
    [Parameter(Mandatory)]$Summary,
    [Parameter(Mandatory)]$Results
  )

  try {
    $rows = @()
    foreach ($r in @($Results)) {
      $serialText = if ([string]::IsNullOrWhiteSpace([string]$r.Serial)) { '-' } else { [string]$r.Serial }
      $resultText = if ([string]::IsNullOrWhiteSpace([string]$r.Result)) { '-' } else { [string]$r.Result }
      $statusText = if ([string]::IsNullOrWhiteSpace([string]$r.Status)) { '-' } else { [string]$r.Status }
      $reasonText = if ([string]::IsNullOrWhiteSpace([string]$r.Reason)) { '-' } else { [string]$r.Reason }

      $isDuplicate = (($statusText -match 'duplicate') -or ($reasonText -match 'already.?assigned|already exists|ztddevicealreadyassigned'))
      if ($isDuplicate) {
        $resultText = 'Blocked'
        $statusText = 'Duplicate'
        if ($reasonText -eq '-' -or $reasonText -match 'ztddevicealreadyassigned') {
          $reasonText = 'Device already exists in Autopilot and was skipped.'
        }
      }

      $rows += [pscustomobject]@{
        Serial = $serialText
        Result = $resultText
        Status = $statusText
        Reason = $reasonText
      }
    }

    $blockedRows = @($rows | Where-Object { $_.Status -eq 'Duplicate' })
    $failedRows  = @($rows | Where-Object { $_.Result -eq 'Failed' -and $_.Status -ne 'Duplicate' })
    $queuedRows  = @($rows | Where-Object { $_.Result -in @('Accepted','Pending') })
    $successRows = @($rows | Where-Object { $_.Result -eq 'Success' })

    $windowTitle = 'Upload Successful'
    $bannerTitle = '! Upload Successful'
    $bannerMessage = 'The device was uploaded successfully.'
    $instruction = 'No action required.'
    $serialHeader = 'Serial Number:'
    $serialList = @($successRows | ForEach-Object { $_.Serial })
    $bannerBg = '#ECFDF3'
    $bannerBorder = '#86EFAC'
    $bannerFg = '#166534'

    if ($failedRows.Count -gt 0) {
      $windowTitle = 'Upload Failed'
      $bannerTitle = '! Upload Failed'
      $bannerMessage = 'The device upload failed.'
      $instruction = 'Verify connectivity and required permissions, then retry the upload.'
      $serialHeader = 'Serial Number:'
      $serialList = @($failedRows | ForEach-Object { $_.Serial })
      $bannerBg = '#FEE2E2'
      $bannerBorder = '#FCA5A5'
      $bannerFg = '#991B1B'
    } elseif ($blockedRows.Count -gt 0 -and ($successRows.Count -gt 0 -or $queuedRows.Count -gt 0)) {
      $windowTitle = 'Upload Completed with Skip'
      $bannerTitle = '! Upload Completed with Skip'
      $bannerMessage = 'The device was uploaded. Existing records were skipped automatically.'
      $instruction = 'No action is required unless you plan to delete and re-import the device.'
      $serialHeader = 'Serial Number:'
      $serialList = @($blockedRows | ForEach-Object { $_.Serial })
      $bannerBg = '#FEF3C7'
      $bannerBorder = '#F59E0B'
      $bannerFg = '#92400E'
    } elseif ($blockedRows.Count -gt 0) {
      $windowTitle = 'Device Already Exists'
      $bannerTitle = '! Upload Blocked'
      $bannerMessage = 'The device already exists in Windows Autopilot.'
      $instruction = 'The device was skipped to prevent duplicate import.'
      $serialHeader = 'Serial Number:'
      $serialList = @($blockedRows | ForEach-Object { $_.Serial })
      $bannerBg = '#FEF3C7'
      $bannerBorder = '#F59E0B'
      $bannerFg = '#92400E'
    } elseif ($queuedRows.Count -gt 0) {
      $windowTitle = 'Upload Queued'
      $bannerTitle = '! Upload Queued'
      $bannerMessage = 'The device import request was accepted and is currently processing in Intune.'
      $instruction = 'Allow processing time, then verify the device status in Intune.'
      $serialHeader = 'Serial Number:'
      $serialList = @($queuedRows | ForEach-Object { $_.Serial })
      $bannerBg = '#DBEAFE'
      $bannerBorder = '#93C5FD'
      $bannerFg = '#1E40AF'
    }

    $serialList = @($serialList | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($serialList.Count -eq 0) { $serialList = @('-') }
    $serialLines = ($serialList | ForEach-Object { "- " + $_ }) -join "`r`n"

    $x = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$windowTitle"
        Width="620" Height="360"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterScreen"
        Background="#F6F8FB"
        FontFamily="Segoe UI"
        FontSize="13">
  <Grid Margin="12">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="12"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="8"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="12"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" x:Name="BannerBorder" Padding="12" CornerRadius="6" BorderThickness="1">
      <StackPanel>
        <TextBlock x:Name="LblBannerTitle" FontSize="18" FontWeight="Bold"/>
        <TextBlock x:Name="LblBannerMessage" Margin="0,4,0,0"/>
      </StackPanel>
    </Border>

    <TextBlock Grid.Row="2" x:Name="LblInstruction" FontSize="14" FontWeight="SemiBold" Foreground="#334155"/>

    <Border Grid.Row="4" Background="#F8FAFD" BorderBrush="#D5DEEA" BorderThickness="1" CornerRadius="6" Padding="10">
      <StackPanel>
        <TextBlock x:Name="LblSerialHeader" FontWeight="SemiBold" Foreground="#0F172A" Margin="0,0,0,6"/>
        <TextBox x:Name="TxtSerials"
                 IsReadOnly="True"
                 BorderThickness="0"
                 Background="#F8FAFD"
                 Foreground="#0F172A"
                 FontFamily="Consolas"
                 FontSize="13"
                 VerticalScrollBarVisibility="Auto"
                 TextWrapping="Wrap"
                 AcceptsReturn="True"/>
      </StackPanel>
    </Border>

    <TextBlock Grid.Row="6" x:Name="LblCounts" Foreground="#64748B" VerticalAlignment="Top"/>

    <Button Grid.Row="8" x:Name="BtnOk" Content="OK"
            Width="110" Height="32"
            HorizontalAlignment="Right"
            Background="#D8E5FF" Foreground="#1E3A6D"
            BorderBrush="#B7CBEF" BorderThickness="1"/>
  </Grid>
</Window>
"@

    $w = [Windows.Markup.XamlReader]::Parse($x)
    $BannerBorder = $w.FindName('BannerBorder')
    $LblBannerTitle = $w.FindName('LblBannerTitle')
    $LblBannerMessage = $w.FindName('LblBannerMessage')
    $LblInstruction = $w.FindName('LblInstruction')
    $LblSerialHeader = $w.FindName('LblSerialHeader')
    $TxtSerials = $w.FindName('TxtSerials')
    $LblCounts = $w.FindName('LblCounts')
    $BtnOk = $w.FindName('BtnOk')

    $BannerBorder.Background = New-Brush $bannerBg
    $BannerBorder.BorderBrush = New-Brush $bannerBorder
    $LblBannerTitle.Text = $bannerTitle
    $LblBannerTitle.Foreground = New-Brush $bannerFg
    $LblBannerMessage.Text = $bannerMessage
    $LblBannerMessage.Foreground = New-Brush $bannerFg

    $LblInstruction.Text = $instruction
    $LblSerialHeader.Text = $serialHeader
    $TxtSerials.Text = $serialLines
    $LblCounts.Text = ("Total: {0}   Success: {1}   Failed: {2}   Duplicate: {3}   Queued: {4}" -f `
      [string]$Summary.Total, [string]$Summary.Success, [string]$Summary.Failed, [string]$Summary.Duplicate, [string]$Summary.Pending)

    $BtnOk.Add_Click({ $w.Close() })
    $w.Add_KeyDown({
      param($sender,$e)
      if ($e.Key -eq [System.Windows.Input.Key]::Escape -or $e.Key -eq [System.Windows.Input.Key]::Enter) {
        $e.Handled = $true
        $sender.Close()
      }
    })
    [void]$w.ShowDialog()
  } catch {
    $detail = [string]$_.Exception.Message
    try { if ($_.ScriptStackTrace) { $detail += (" | " + [string]$_.ScriptStackTrace) } } catch { }
    Add-Log ("Results window error: " + $detail) "ERROR"
  }
}
#endregion ============================== DEVICE / ENVIRONMENT CHECKS =================================

#region ============================== GRAPH STATE + UI SYNC =========================================
$script:GraphConnected = $false
$script:GraphAccount   = $null
$script:GraphTenantId  = $null
$script:GraphConnecting = $false
$script:CachedGraphAccessToken = $null
$script:MgGraphModuleReady = $false
$script:MsalModuleReady = $false

# Function: Update-MainGraphUI
# Applies Graph state to status labels, pills, and buttons.
function Update-MainGraphUI {
  if ($FooterCenter) { $FooterCenter.Text = "Version $AppVersion" }

  if ($BtnGraphConnect) {
    if ($script:GraphConnecting) { $BtnGraphConnect.Content = 'Connecting...' }
    else                         { $BtnGraphConnect.Content = 'Connect' }
  }

  if ($script:GraphConnecting) {
    $LblStatus.Text = 'Connecting...'
    $LblStatus.Foreground = New-Brush '#92400E'
    $LblUser.Text = '-'
    $LblTenant.Text = '-'
  }
  elseif ($script:GraphConnected) {
    $LblStatus.Text = 'Connected'
    $LblStatus.Foreground = New-Brush '#0A8A0A'
    if ($script:GraphAccount)  { $LblUser.Text   = $script:GraphAccount } else { $LblUser.Text = '-' }
    if ($script:GraphTenantId) { $LblTenant.Text = $script:GraphTenantId } else { $LblTenant.Text = '-' }
  } else {
    $LblStatus.Text  = 'Not Connected'
    $LblStatus.Foreground = New-Brush 'Black'
    $LblUser.Text   = '-'
    $LblTenant.Text = '-'
  }

  if ($SideGraphTxt -and $SideGraphPill) {
    if ($script:GraphConnecting) {
      $SideGraphTxt.Text = 'Connecting...'
      $SideGraphPill.Background = New-Brush '#FEF3C7'
      $SideGraphTxt.Foreground  = New-Brush '#92400E'
    }
    elseif ($script:GraphConnected) {
      $SideGraphTxt.Text = 'Connected'
      $SideGraphPill.Background = New-Brush '#ECFDF3'
      $SideGraphTxt.Foreground  = New-Brush '#166534'
    } else {
      $SideGraphTxt.Text = 'Not Connected'
      $SideGraphPill.Background = New-Brush '#EEF2FF'
      $SideGraphTxt.Foreground  = New-Brush '#1D4ED8'
    }
  }

  if ($BtnGraphConnect -and $BtnGraphDisconnect) {
    if ($script:GraphConnecting) {
      $BtnGraphConnect.IsEnabled = $false
      $BtnGraphDisconnect.IsEnabled = $false
    } else {
      $BtnGraphConnect.IsEnabled = (-not $script:GraphConnected)
      $BtnGraphDisconnect.IsEnabled = [bool]$script:GraphConnected
    }
  }
}
#endregion ============================== GRAPH STATE + UI SYNC =======================================

#region ============================== BACKGROUND PIPELINE (RUNSPACE + TIMER) ========================
$iss  = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
$pool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 1, $iss, $Host)
$pool.ApartmentState = [System.Threading.ApartmentState]::STA
$pool.Open()

$script:PS            = $null
$script:Async         = $null
$script:CurrentAction = ''
$script:LastFailedRecords = @()
$script:UploadProgressActive = $false
$script:AppCleanupDone = $false

function Invoke-AppCleanup {
  if ($script:AppCleanupDone) { return }
  $script:AppCleanupDone = $true

  try { Stop-UploadProgress } catch { }
  try { if ($Timer) { $Timer.Stop() } } catch { }

  try {
    if ($script:PS -and $script:Async -and -not $script:Async.IsCompleted) {
      $script:PS.Stop()
    }
  } catch { }

  try {
    if ($script:PS) { $script:PS.Dispose() }
  } catch { }
  $script:PS = $null
  $script:Async = $null

  try {
    if ($pool) {
      $pool.Close()
      $pool.Dispose()
    }
  } catch { }

  try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
  } catch { }
  $script:CachedGraphAccessToken = $null

  try {
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  } catch { }
}

$Timer = New-Object Windows.Threading.DispatcherTimer
$Timer.Interval = [TimeSpan]::FromMilliseconds(300)
$Timer.Add_Tick({
  try {
    if ($script:UploadProgressActive -and $script:Async -and -not $script:Async.IsCompleted -and ($script:CurrentAction -in @('enroll','retry'))) {
      Step-UploadProgress
    }

    if ($script:Async -and $script:Async.IsCompleted) {
      $out = $script:PS.EndInvoke($script:Async)
      try { $script:PS.Dispose() } catch { }
      $script:PS    = $null
      $script:Async = $null
      $Timer.Stop()

      switch ($script:CurrentAction) {
        'collect' { $BtnCollectHash.IsEnabled = $true }
        'enroll'  {
          $BtnEnroll.IsEnabled = $true
          if ($BtnRetryFailed) { $BtnRetryFailed.IsEnabled = ($script:LastFailedRecords.Count -gt 0) }
          Stop-UploadProgress
        }
        'retry'   {
          $BtnEnroll.IsEnabled = $true
          if ($BtnRetryFailed) { $BtnRetryFailed.IsEnabled = ($script:LastFailedRecords.Count -gt 0) }
          Stop-UploadProgress
        }
        'graphauth' {
          $script:GraphConnecting = $false
          if ($BtnGraphConnect) { $BtnGraphConnect.IsEnabled = $true }
          Update-MainGraphUI
        }
      }

      $txt = ($out | Out-String).Trim()
      if (-not $txt) { return }
      $obj = $null
      try { $obj = $txt | ConvertFrom-Json } catch { return }

      if ($obj.Action -eq 'collect') {
        if ($obj.Success) { Add-Log ("Hash collected: " + $obj.Path) 'SUCCESS' }
        else              { Add-Log ("Collect error: " + $obj.Error) 'ERROR' }
      }
      elseif ($obj.Action -eq 'enroll') {
        if ($obj.CsvCreated -and $obj.CsvPath) { Add-Log ("CSV created: " + $obj.CsvPath) 'INFO' }
        if ($obj.Summary) {
          Add-Log ("Upload summary: total=" + $obj.Summary.Total + ", success=" + $obj.Summary.Success + ", failed=" + $obj.Summary.Failed + ", duplicate=" + $obj.Summary.Duplicate + ", queued=" + $obj.Summary.Pending) 'INFO'
        }
        if ($obj.Results) {
          foreach ($r in @($obj.Results)) {
            if ($r.Result -eq 'Failed')  { Add-Log ("Failed [" + $r.Serial + "]: " + $r.Reason) 'ERROR' }
            if ($r.Status -eq 'Duplicate' -or $r.Result -eq 'Skipped') {
              Add-Log ("Duplicate skipped [" + $r.Serial + "]: already exists in Autopilot.") 'WARN'
            }
          }
        }

        $script:LastFailedRecords = @()
        if ($obj.FailedRecords) { $script:LastFailedRecords = @($obj.FailedRecords) }
        if ($BtnRetryFailed) { $BtnRetryFailed.IsEnabled = ($script:LastFailedRecords.Count -gt 0) }

        if ($obj.Success) { Add-Log ("Uploaded " + $obj.Uploaded + " record(s) to Autopilot.") 'SUCCESS' }
        else {
          if ($obj.Error) { Add-Log ("Upload error: " + $obj.Error) 'ERROR' }
          else            { Add-Log "Upload completed with warnings or failures." 'WARN' }
        }

        if ($obj.Results -and $obj.Summary) {
          Show-EnrollResultsWindow -Summary $obj.Summary -Results @($obj.Results)
        }
      }
      elseif ($obj.Action -eq 'graphauth') {
        if ($obj.Success -and $obj.AccessToken) {
          try {
            $rawToken = [string]$obj.AccessToken
            $rawToken = $rawToken.Trim()
            $secureToken = ConvertTo-SecureString -String $rawToken -AsPlainText -Force
            Connect-MgGraph -AccessToken $secureToken -NoWelcome -ErrorAction Stop | Out-Null
            if (Sync-GraphContext -Silent) {
              $script:CachedGraphAccessToken = $rawToken
              Add-Log "Connected to Graph (Browser Interactive)." 'SUCCESS'
            } else {
              $script:CachedGraphAccessToken = $null
              Add-Log "Graph connection did not complete." 'WARN'
            }
          } catch {
            $script:CachedGraphAccessToken = $null
            Add-Log ("Graph connect error: " + $_.Exception.Message) 'ERROR'
          }
        } else {
          if ($obj.Error) { Add-Log ("Graph connect error: " + $obj.Error) 'ERROR' }
          else            { Add-Log "Graph connect failed." 'ERROR' }
        }
        Update-MainGraphUI
      }
    }
  } catch {
    Add-Log ('Async error: ' + $_.Exception.Message) 'ERROR'
    if ($script:CurrentAction -in @('enroll','retry')) { Stop-UploadProgress }
    if ($script:CurrentAction -eq 'graphauth') {
      $script:GraphConnecting = $false
      try { Update-MainGraphUI } catch { }
    }
    try { $script:PS.Dispose() } catch { }
    $script:PS = $null; $script:Async = $null
    $Timer.Stop()
  }
})
#endregion ============================== BACKGROUND PIPELINE (RUNSPACE + TIMER) ======================

#region ============================== WORKER SCRIPTS ================================================
# 0) Graph token worker (cache-first, interactive fallback)
$GraphTokenWorker = @'
param($Scopes)
try{
  $tmpTokenFile = [System.IO.Path]::GetTempFileName()
  $tmpErrFile   = [System.IO.Path]::GetTempFileName()
  $safeTmpPath  = $tmpTokenFile.Replace("'", "''")
  $safeErrPath  = $tmpErrFile.Replace("'", "''")
  $msalModulePath = $null
  try {
    $msalModulePath = (Get-Module -ListAvailable -Name MSAL.PS | Sort-Object Version -Descending | Select-Object -First 1).Path
  } catch { $msalModulePath = $null }
  $safeMsalModulePath = ''
  if ($msalModulePath) { $safeMsalModulePath = [string]$msalModulePath.Replace("'", "''") }

  $scopeArrayLiteral = (@($Scopes) | ForEach-Object { "'" + ([string]$_).Replace("'", "''") + "'" }) -join ','
  if([string]::IsNullOrWhiteSpace($scopeArrayLiteral)){ $scopeArrayLiteral = "'User.Read'" }

  $childScript = @"
`$ErrorActionPreference = 'Stop'
try{
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
  if(-not (Get-Command -Name Import-PowerShellDataFile -ErrorAction SilentlyContinue)){
    function Import-PowerShellDataFile {
      param([Parameter(Mandatory=`$true)][string]`$Path)
      if(-not (Test-Path -LiteralPath `$Path)){
        throw ("Import-PowerShellDataFile fallback: file not found: " + `$Path)
      }
      `$raw = Get-Content -LiteralPath `$Path -Raw -ErrorAction Stop
      `$data = & ([scriptblock]::Create(`$raw))
      if(-not (`$data -is [hashtable])){
        throw ("Import-PowerShellDataFile fallback: invalid data file: " + `$Path)
      }
      return `$data
    }
  }
  `$msalModulePath = '$safeMsalModulePath'
  if(-not [string]::IsNullOrWhiteSpace(`$msalModulePath) -and (Test-Path -LiteralPath `$msalModulePath)){
    Import-Module -Name `$msalModulePath -ErrorAction Stop
  } else {
    Import-Module MSAL.PS -ErrorAction Stop
  }

  `$scopes = @($scopeArrayLiteral)
  `$tokenScopes = @(`$scopes | ForEach-Object { [string]`$_ })
  if(`$tokenScopes -notcontains 'openid')         { `$tokenScopes += 'openid' }
  if(`$tokenScopes -notcontains 'profile')        { `$tokenScopes += 'profile' }
  if(`$tokenScopes -notcontains 'offline_access') { `$tokenScopes += 'offline_access' }

  # Microsoft Graph PowerShell public client ID.
  `$publicClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
  `$msalToken = `$null
  try {
    # Fast path from token cache (no browser prompt).
    `$msalToken = Get-MsalToken -ClientId `$publicClientId -Scopes `$tokenScopes -Silent -ErrorAction Stop
  } catch {
    `$msalToken = `$null
  }
  if(-not `$msalToken -or -not `$msalToken.AccessToken){
    `$msalToken = Get-MsalToken -ClientId `$publicClientId -Scopes `$tokenScopes -Interactive -ErrorAction Stop
  }
  if(-not `$msalToken -or -not `$msalToken.AccessToken){
    throw 'Could not acquire access token from browser interactive sign-in.'
  }

  `$utf8NoBom = New-Object System.Text.UTF8Encoding(`$false)
  [System.IO.File]::WriteAllText('$safeTmpPath', ([string]`$msalToken.AccessToken), `$utf8NoBom)
  exit 0
}catch{
  try {
    `$errMsg = `$_.Exception.Message
    if(`$_.ScriptStackTrace){ `$errMsg += (' | ' + `$_.ScriptStackTrace) }
    Set-Content -LiteralPath '$safeErrPath' -Value `$errMsg -Encoding UTF8 -Force
  } catch { }
  exit 1
}
"@

  $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childScript))
  $psExe = Join-Path $PSHOME 'powershell.exe'
  if(-not (Test-Path -LiteralPath $psExe)){ $psExe = 'powershell.exe' }

  $proc = Start-Process -FilePath $psExe `
    -ArgumentList @('-NoProfile','-NoLogo','-ExecutionPolicy','Bypass','-STA','-EncodedCommand',$encodedCommand) `
    -WindowStyle Minimized -Wait -PassThru

  if($proc.ExitCode -ne 0){
    $childErr = ''
    try { $childErr = Get-Content -LiteralPath $tmpErrFile -Raw -ErrorAction Stop } catch { $childErr = '' }
    if([string]::IsNullOrWhiteSpace($childErr)){
      throw ("Browser sign-in process failed with exit code " + $proc.ExitCode + ".")
    } else {
      throw ("Browser sign-in process failed: " + $childErr.Trim())
    }
  }

  $accessToken = ''
  try { $accessToken = Get-Content -LiteralPath $tmpTokenFile -Raw -ErrorAction Stop } catch { $accessToken = '' }
  if($accessToken){ $accessToken = $accessToken.Trim() }
  try { Remove-Item -LiteralPath $tmpTokenFile -Force -ErrorAction SilentlyContinue } catch { }
  try { Remove-Item -LiteralPath $tmpErrFile -Force -ErrorAction SilentlyContinue } catch { }

  if([string]::IsNullOrWhiteSpace($accessToken)){
    throw "Could not acquire access token from browser interactive sign-in."
  }

  [pscustomobject]@{
    Action='graphauth'
    Success=$true
    AccessToken=[string]$accessToken
    Error=$null
  } | ConvertTo-Json -Compress
}catch{
  try { if($tmpTokenFile){ Remove-Item -LiteralPath $tmpTokenFile -Force -ErrorAction SilentlyContinue } } catch { }
  try { if($tmpErrFile){ Remove-Item -LiteralPath $tmpErrFile -Force -ErrorAction SilentlyContinue } } catch { }
  [pscustomobject]@{
    Action='graphauth'
    Success=$false
    AccessToken=$null
    Error=$_.Exception.Message
  } | ConvertTo-Json -Compress
}
'@

# 1) Collect HWID to CSV
$CollectHWIDWorker = @'
param($OutFolder)
try{
  if([string]::IsNullOrWhiteSpace($OutFolder)){ throw "Please select a folder." }
  if(-not (Test-Path -LiteralPath $OutFolder)){ [void](New-Item -ItemType Directory -Path $OutFolder -Force) }

  $serial = (Get-CimInstance Win32_BIOS -ErrorAction Stop).SerialNumber
  $prodId = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductId).ProductId
  $devMap = Get-CimInstance -Namespace root\cimv2\mdm\dmmap -Class MDM_DevDetail_Ext01 -ErrorAction Stop
  $hw     = $devMap.DeviceHardwareData
  if(-not $serial -or -not $hw){ throw "Could not read Serial/Hardware Hash. Run as admin and ensure device supports Autopilot hash." }

  $ts   = (Get-Date).ToString("yyyyMMdd_HHmmss")
  $file = Join-Path $OutFolder ("AutopilotHWID_{0}_{1}.csv" -f $env:COMPUTERNAME,$ts)

  Set-Content -LiteralPath $file -Value 'Device Serial Number,Windows Product ID,Hardware Hash' -Encoding UTF8
  Add-Content -LiteralPath $file -Value ('"'+$serial+'","'+$prodId+'","'+$hw+'"' ) -Encoding UTF8

  [pscustomobject]@{Action='collect';Success=$true;Path=$file;Error=$null} | ConvertTo-Json -Compress
}catch{
  [pscustomobject]@{Action='collect';Success=$false;Path=$null;Error=$_.Exception.Message} | ConvertTo-Json -Compress
}
'@

# 2) Upload/Import to Autopilot (auto-create CSV when no valid input file is provided)
$EnrollWorker = @'
param($Path,$GroupTag,$AssignedUser,$AssignedName,$HwIdFolder,$RetryJson)
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
try{
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
  try {
    $sysProxy = [System.Net.WebRequest]::GetSystemWebProxy()
    if ($sysProxy) {
      $sysProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
      [System.Net.WebRequest]::DefaultWebProxy = $sysProxy
    }
  } catch { }

  $ctx=$null; try{$ctx=Get-MgContext}catch{}
  if(-not $ctx -or (-not $ctx.Account -and -not $ctx.ClientId)){ throw "Not connected to Graph." }

  function New-ResultRow {
    param($Serial,$Result,$Status,$Reason)
    return [pscustomobject]@{
      Serial = [string]$Serial
      Result = [string]$Result
      Status = [string]$Status
      Reason = [string]$Reason
    }
  }

  function Get-GraphErrorText {
    param($ErrRecord)

    $parts = @()
    try {
      if ($ErrRecord -and $ErrRecord.Exception -and $ErrRecord.Exception.Message) {
        $parts += [string]$ErrRecord.Exception.Message
      }
    } catch { }

    try {
      if ($ErrRecord -and $ErrRecord.Exception -and $ErrRecord.Exception.InnerException -and $ErrRecord.Exception.InnerException.Message) {
        $parts += [string]$ErrRecord.Exception.InnerException.Message
      }
    } catch { }

    try {
      if ($ErrRecord -and $ErrRecord.ErrorDetails -and $ErrRecord.ErrorDetails.Message) {
        $detail = [string]$ErrRecord.ErrorDetails.Message
        try {
          $j = $detail | ConvertFrom-Json -ErrorAction Stop
          if ($j -and $j.error -and $j.error.message) { $parts += [string]$j.error.message }
          else { $parts += $detail }
        } catch {
          $parts += $detail
        }
      }
    } catch { }

    $clean = @($parts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($clean.Count -gt 0) { return ($clean -join ' | ') }
    return 'Unknown Graph error.'
  }

  function Normalize-Record {
    param($r,$DefaultGroupTag,$DefaultUser,$DefaultName)
    $serial = $null
    $prod   = $null
    $hash   = $null
    if($r.PSObject.Properties.Match('serialNumber').Count -gt 0)            { $serial = $r.serialNumber }
    elseif($r.PSObject.Properties.Match('Device Serial Number').Count -gt 0) { $serial = $r.'Device Serial Number' }
    if($r.PSObject.Properties.Match('productKey').Count -gt 0)               { $prod = $r.productKey }
    elseif($r.PSObject.Properties.Match('Windows Product ID').Count -gt 0)   { $prod = $r.'Windows Product ID' }
    if($r.PSObject.Properties.Match('hardwareIdentifier').Count -gt 0)        { $hash = $r.hardwareIdentifier }
    elseif($r.PSObject.Properties.Match('Hardware Hash').Count -gt 0)         { $hash = $r.'Hardware Hash' }

    $gt = if($r.PSObject.Properties.Match('groupTag').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$r.groupTag)) { [string]$r.groupTag } else { [string]$DefaultGroupTag }
    $au = if($r.PSObject.Properties.Match('assignedUserPrincipalName').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$r.assignedUserPrincipalName)) { [string]$r.assignedUserPrincipalName } else { [string]$DefaultUser }
    $an = if($r.PSObject.Properties.Match('assignedComputerName').Count -gt 0 -and -not [string]::IsNullOrWhiteSpace([string]$r.assignedComputerName)) { [string]$r.assignedComputerName } else { [string]$DefaultName }

    return [pscustomobject]@{
      serialNumber = [string]$serial
      productKey = [string]$prod
      hardwareIdentifier = [string]$hash
      groupTag = $gt
      assignedUserPrincipalName = $au
      assignedComputerName = $an
    }
  }

  $records=@(); $csvCreated=$false; $csvPath=$null
  $results=@(); $failedRecords=@()

  try {
    Invoke-MgGraphRequest -Method GET -Uri '/v1.0/organization?$top=1' -ErrorAction Stop | Out-Null
  } catch {
    $checkReason = Get-GraphErrorText -ErrRecord $_
    if ($checkReason -match 'Invalid Hostname|request hostname is invalid') {
      $checkReason += ' | Check proxy/WinHTTP settings on this device.'
    }
    throw ("Graph connectivity check failed: " + $checkReason)
  }

  if(-not [string]::IsNullOrWhiteSpace($RetryJson)){
    $retryList = $null
    try { $retryList = $RetryJson | ConvertFrom-Json } catch { throw "Retry payload is not valid JSON." }
    foreach($r in @($retryList)){
      $records += Normalize-Record -r $r -DefaultGroupTag $GroupTag -DefaultUser $AssignedUser -DefaultName $AssignedName
    }
  }
  elseif([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path) -or (Test-Path -LiteralPath $Path -PathType Container)){
    if(-not (Test-Path -LiteralPath $HwIdFolder)){ New-Item -ItemType Directory -Path $HwIdFolder -Force | Out-Null }
    $serial = (Get-CimInstance Win32_BIOS -ErrorAction Stop).SerialNumber
    $prodId = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductId).ProductId
    $devMap = Get-CimInstance -Namespace root\cimv2\mdm\dmmap -Class MDM_DevDetail_Ext01 -ErrorAction Stop
    $hw     = $devMap.DeviceHardwareData
    if(-not $serial -or -not $hw){ throw "Could not read Serial/Hardware Hash." }

    $ts      = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $csvPath = Join-Path $HwIdFolder ("AutopilotHWID_{0}_{1}.csv" -f $env:COMPUTERNAME,$ts)
    Set-Content -LiteralPath $csvPath -Value 'Device Serial Number,Windows Product ID,Hardware Hash' -Encoding UTF8
    Add-Content -LiteralPath $csvPath -Value ('"'+$serial+'","'+$prodId+'","'+$hw+'"' ) -Encoding UTF8
    $csvCreated=$true

    $records += [pscustomobject]@{
      serialNumber=$serial; productKey=$prodId; hardwareIdentifier=$hw;
      groupTag=$GroupTag; assignedUserPrincipalName=$AssignedUser; assignedComputerName=$AssignedName
    }
  } else {
    if($Path.ToLower().EndsWith(".json")){
      $json=Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json
      foreach($r in @($json)){ $records += Normalize-Record -r $r -DefaultGroupTag $GroupTag -DefaultUser $AssignedUser -DefaultName $AssignedName }
    } else {
      $csv=Import-Csv -LiteralPath $Path
      foreach($r in $csv){
        $records += Normalize-Record -r $r -DefaultGroupTag $GroupTag -DefaultUser $AssignedUser -DefaultName $AssignedName
      }
    }
  }
  if($records.Count -eq 0){ throw "No records to upload." }

  $uploaded=0
  foreach($rec in $records){
    $serial = [string]$rec.serialNumber
    $prodId = [string]$rec.productKey
    $hash   = [string]$rec.hardwareIdentifier
    if ($serial) { $serial = $serial.Trim() }
    if ($prodId) { $prodId = $prodId.Trim() }
    if ($hash)   { $hash   = $hash.Trim() }
    if([string]::IsNullOrWhiteSpace($serial) -or [string]::IsNullOrWhiteSpace($hash)){
      $results += New-ResultRow -Serial $serial -Result 'Failed' -Status 'Validation' -Reason 'Missing serial number or hardware hash.'
      $failedRecords += $rec
      continue
    }

    # Duplicate check by serial in active Autopilot devices only.
    # Imported identities can be transient during tenant sync and should not hard-block here.
    $isDuplicate = $false
    $safeSerial = $serial.Replace("'","''")
    $filter = "serialNumber eq '$safeSerial'"
    $dupFilter = [System.Uri]::EscapeDataString($filter)

    try {
      $dupUriWindows = "/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=$dupFilter&`$top=1"
      $dupResWindows = Invoke-MgGraphRequest -Method GET -Uri $dupUriWindows -ErrorAction Stop
      if($dupResWindows.value -and $dupResWindows.value.Count -gt 0){ $isDuplicate = $true }
    } catch { }

    if($isDuplicate){
      $results += New-ResultRow -Serial $serial -Result 'Skipped' -Status 'Duplicate' -Reason 'Device already exists in Autopilot and was skipped.'
      continue
    }

    $uri='/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities'
    $bodyMap = @{
      serialNumber       = $serial
      hardwareIdentifier = $hash
    }
    if (-not [string]::IsNullOrWhiteSpace($prodId)) { $bodyMap.productKey = $prodId }
    if (-not [string]::IsNullOrWhiteSpace([string]$rec.groupTag)) { $bodyMap.groupTag = [string]$rec.groupTag }
    if (-not [string]::IsNullOrWhiteSpace([string]$rec.assignedUserPrincipalName)) { $bodyMap.assignedUserPrincipalName = [string]$rec.assignedUserPrincipalName }
    $body = $bodyMap | ConvertTo-Json -Depth 5

    try {
      $postRes = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json' -ErrorAction Stop
      $uploaded++

      $importId = $null
      if($postRes -and $postRes.id){ $importId = [string]$postRes.id }

      $finalStatus = 'pending'
      $reason = 'Still pending in Graph.'
      $lastPollError = $null
      $safeSerial = $serial.Replace("'","''")
      $filter = "serialNumber eq '$safeSerial'"
      $checkUri = "/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=$([System.Uri]::EscapeDataString($filter))&`$top=1"
      $maxPollAttempts = 20
      $pollDelaySeconds = 15

      for($i=1; $i -le $maxPollAttempts; $i++){
        Start-Sleep -Seconds $pollDelaySeconds

        # Poll imported identity state when importId is available.
        if($importId){
          try {
            $item = Invoke-MgGraphRequest -Method GET -Uri ("/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities/" + $importId) -ErrorAction Stop
            $state = $item.state
            if($state -and $state.deviceImportStatus){ $finalStatus = [string]$state.deviceImportStatus }
            if([string]::IsNullOrWhiteSpace($finalStatus)){ $finalStatus = 'pending' }
            if($finalStatus -eq 'unknown'){ $finalStatus = 'pending' }

            $errName = ''
            $errCode = ''
            if($state){
              if($state.deviceErrorName){ $errName = [string]$state.deviceErrorName }
              if($state.deviceErrorCode -ne $null){ $errCode = [string]$state.deviceErrorCode }
            }

            if($finalStatus -match 'complete|completed|success'){
              $reason = 'Imported successfully.'
              break
            }
            if($finalStatus -match 'error|failed'){
              if($errName -or $errCode){ $reason = ("{0} {1}" -f $errName, $errCode).Trim() }
              else { $reason = 'Import failed.' }
              break
            }
          } catch {
            $lastPollError = $_.Exception.Message
          }
        }

        # Poll Intune Autopilot devices every 15 seconds and mark success as soon as device appears.
        try {
          $checkRes = Invoke-MgGraphRequest -Method GET -Uri $checkUri -ErrorAction Stop
          if($checkRes.value -and $checkRes.value.Count -gt 0){
            $finalStatus = 'complete'
            $reason = ("Imported successfully (found in Autopilot devices after {0} seconds)." -f ($i * $pollDelaySeconds))
            break
          }
        } catch {
          $lastPollError = $_.Exception.Message
        }

        $reason = ("Still queued in Intune (attempt {0}/{1})." -f $i, $maxPollAttempts)
      }

      if($finalStatus -match 'pending|unknown'){
        if($importId){
          $reason = ("Still queued in Intune. ImportId: " + $importId)
        } else {
          $reason = "Still queued in Intune."
        }
        if($lastPollError){ $reason += (" | Last poll error: " + $lastPollError) }
      }

      if($finalStatus -match 'complete|completed|success'){
        $results += New-ResultRow -Serial $serial -Result 'Success' -Status 'Complete' -Reason $reason
      }
      elseif($finalStatus -match 'error|failed'){
        $results += New-ResultRow -Serial $serial -Result 'Failed' -Status $finalStatus -Reason $reason
        $failedRecords += $rec
      }
      else {
        $results += New-ResultRow -Serial $serial -Result 'Accepted' -Status 'Queued' -Reason $reason
      }
    } catch {
      $reason = Get-GraphErrorText -ErrRecord $_
      if ($reason -match 'Invalid Hostname|request hostname is invalid') {
        $reason += ' | Check proxy/WinHTTP settings on this device.'
      }
      if ($reason -match 'ZtdDeviceAlreadyAssigned|already assigned') {
        $results += New-ResultRow -Serial $serial -Result 'Skipped' -Status 'Duplicate' -Reason 'Device already exists in Autopilot and was skipped.'
        continue
      }
      $results += New-ResultRow -Serial $serial -Result 'Failed' -Status 'UploadError' -Reason $reason
      $failedRecords += $rec
    }
  }

  $summary = [pscustomobject]@{
    Total = $results.Count
    Success = (@($results | Where-Object { $_.Result -eq 'Success' })).Count
    Failed = (@($results | Where-Object { $_.Result -eq 'Failed' -and $_.Status -ne 'Duplicate' })).Count
    Duplicate = (@($results | Where-Object { $_.Status -eq 'Duplicate' -or $_.Result -eq 'Skipped' })).Count
    Pending = (@($results | Where-Object { $_.Result -in @('Pending','Accepted') })).Count
  }
  $overallSuccess = ($summary.Failed -eq 0 -and ($summary.Success -gt 0 -or $summary.Pending -gt 0 -or $summary.Duplicate -gt 0))

  [pscustomobject]@{
    Action='enroll'
    Success=$overallSuccess
    Uploaded=$uploaded
    CsvCreated=$csvCreated
    CsvPath=$csvPath
    Summary=$summary
    Results=$results
    FailedRecords=$failedRecords
    Error=$null
  } | ConvertTo-Json -Depth 8 -Compress
}catch{
  [pscustomobject]@{
    Action='enroll'
    Success=$false
    Uploaded=0
    CsvCreated=$false
    CsvPath=$null
    Summary=[pscustomobject]@{Total=0;Success=0;Failed=0;Duplicate=0;Pending=0}
    Results=@()
    FailedRecords=@()
    Error=$_.Exception.Message
  } | ConvertTo-Json -Depth 8 -Compress
}
'@
#endregion ============================== WORKER SCRIPTS ==============================================

#region ============================== MAIN WINDOW UI ================================================
$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Autopilot Assistant"
        Width="1180" Height="800"
        WindowStartupLocation="CenterScreen"
        Background="#F6F8FB"
        FontFamily="Segoe UI"
        FontSize="13"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True">

  <Window.Resources>

    <!-- Shadows -->
    <DropShadowEffect x:Key="ShadowBlue"  BlurRadius="10" ShadowDepth="0" Opacity="0.55" Color="#9FAEF7"/>
    <DropShadowEffect x:Key="ShadowGray"  BlurRadius="10" ShadowDepth="0" Opacity="0.25" Color="#9CA3AF"/>

    <!-- Buttons -->
    <Style x:Key="BtnBase" TargetType="Button">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Padding" Value="12,0"/>
      <Setter Property="Height" Value="34"/>
      <Setter Property="MinWidth" Value="110"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Style.Triggers>
        <Trigger Property="IsEnabled" Value="False">
          <Setter Property="Effect" Value="{x:Null}"/>
          <Setter Property="Background" Value="#ECEFF3"/>
          <Setter Property="Foreground" Value="#9CA3AF"/>
          <Setter Property="Opacity" Value="0.75"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="BtnPrimary" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#9FAEF7"/>
      <Setter Property="Foreground" Value="#1F2D3A"/>
      <Setter Property="Effect" Value="{StaticResource ShadowBlue}"/>
    </Style>

    <Style x:Key="BtnNeutral" TargetType="Button" BasedOn="{StaticResource BtnBase}">
      <Setter Property="Background" Value="#D8E5FF"/>
      <Setter Property="Foreground" Value="#1E3A6D"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush" Value="#B7CBEF"/>
      <Setter Property="Effect" Value="{StaticResource ShadowGray}"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#C7D9FB"/>
          <Setter Property="BorderBrush" Value="#9FB9E9"/>
        </Trigger>
        <Trigger Property="IsPressed" Value="True">
          <Setter Property="Background" Value="#B8CDF5"/>
          <Setter Property="BorderBrush" Value="#8FADE5"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="BtnSmall" TargetType="Button" BasedOn="{StaticResource BtnNeutral}">
      <Setter Property="Height" Value="28"/>
      <Setter Property="MinWidth" Value="90"/>
      <Setter Property="Padding" Value="10,0"/>
      <Setter Property="Background" Value="#DCE8FF"/>
      <Setter Property="Foreground" Value="#1E3A6D"/>
    </Style>

    <!-- TextBox base -->
    <Style TargetType="TextBox">
      <Setter Property="Height" Value="28"/>
      <Setter Property="Padding" Value="6,3"/>
    </Style>

    <!-- Sidebar cards -->
    <SolidColorBrush x:Key="SidebarCardBackground" Color="#F9FBFF"/>
    <SolidColorBrush x:Key="SidebarCardBorder"     Color="#E4E9F0"/>
    <Style x:Key="SidebarCard" TargetType="Border">
      <Setter Property="Background" Value="{StaticResource SidebarCardBackground}"/>
      <Setter Property="BorderBrush" Value="{StaticResource SidebarCardBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="5"/>
      <Setter Property="Padding" Value="12"/>
      <Setter Property="Margin" Value="12,10,12,0"/>
    </Style>

    <Style x:Key="SidebarTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#0F172A"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>

    <Style x:Key="SidebarText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#4B5563"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>

    <!-- Device information value pills -->
    <Style x:Key="DeviceValuePill" TargetType="Border">
      <Setter Property="Background" Value="#EAF3FF"/>
      <Setter Property="BorderBrush" Value="#e2eeff"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="0"/>
      <Setter Property="Padding" Value="8,2"/>
      <Setter Property="Margin" Value="1"/>
    </Style>
    <Style x:Key="DeviceValueText" TargetType="TextBlock">
      <Setter Property="Foreground" Value="#0F3D91"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
    </Style>

  </Window.Resources>

  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="320"/>
      <ColumnDefinition Width="*"/>
    </Grid.ColumnDefinitions>

    <!-- Sidebar -->
    <Border Grid.Column="0" Background="White" BorderBrush="#E6EBF4" BorderThickness="0,0,1,0">
      <DockPanel LastChildFill="True">

        <!-- App Header -->
        <StackPanel DockPanel.Dock="Top" Margin="18,18,18,12">
          <StackPanel Orientation="Horizontal">
            <Border Width="36" Height="36" Background="#9AB8FF" CornerRadius="6">
              <TextBlock Text="A" Foreground="#1F2D3A" FontSize="18" FontWeight="Bold"
                         VerticalAlignment="Center" HorizontalAlignment="Center"/>
            </Border>
            <StackPanel Margin="10,0,0,0">
              <TextBlock Text="Autopilot Assistant" FontSize="16" FontWeight="SemiBold" Foreground="#1F2D3A"/>
              <TextBlock Text="Provisioning Toolkit" FontSize="11" Foreground="#5F6B7A"/>
            </StackPanel>
          </StackPanel>
        </StackPanel>

        <!-- Footer -->
        <Border DockPanel.Dock="Bottom" BorderBrush="#E6EBF4" BorderThickness="0,1,0,0" Padding="14" Background="#FFFFFF">
          <StackPanel>
            <TextBlock x:Name="FooterLeft" Text="Autopilot Assistant" FontSize="13" FontWeight="Bold" Foreground="#1F2D3A"/>
            <TextBlock x:Name="FooterCenter" Text="Version" FontSize="11" Foreground="#5F6B7A" Margin="0,4,0,0"/>
            <TextBlock x:Name="FooterRight" FontSize="11" Foreground="#7C8BA1" Margin="0,8,0,0">
              <Run x:Name="FooterCopyRun" Text="(c) 2025 "/>
              <Hyperlink x:Name="FooterLink" NavigateUri="https://www.linkedin.com/in/mabdulkadr/">Mohammad Omar</Hyperlink>
            </TextBlock>
          </StackPanel>
        </Border>

        <!-- Nav -->
        <StackPanel DockPanel.Dock="Top" Margin="8,8">
          <TextBlock Text="TOOLS" Margin="14,10,0,6" FontSize="11" FontWeight="SemiBold" Foreground="#7C8BA1"/>
          <Button Content="Autopilot" FontWeight="SemiBold" Height="38" Margin="6" Padding="12,0"
                  ToolTip="Autopilot section (current)"
                  HorizontalContentAlignment="Left"
                  Background="#D8E2F4" Foreground="#1F2D3A" BorderThickness="0"/>
        </StackPanel>

        <!-- Session + About -->
        <Grid>
          <StackPanel VerticalAlignment="Bottom">

            <!-- Graph status pill -->
            <Border Style="{StaticResource SidebarCard}" Margin="12,0,12,8">
              <StackPanel>
                <TextBlock Text="Microsoft Graph" Style="{StaticResource SidebarTitle}"/>
                <Border x:Name="SideGraphPill" Background="#EEF2FF" Padding="6,2" CornerRadius="4" Margin="0,0,0,4">
                  <TextBlock x:Name="SideGraphTxt" Text="Not Connected" Foreground="#1D4ED8" FontWeight="SemiBold"/>
                </Border>

                <Grid Margin="0,4,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>

                  <TextBlock Grid.Row="0" Grid.Column="0" Text="Status:" Style="{StaticResource SidebarText}" FontWeight="SemiBold" Foreground="#111827" Margin="0,0,6,4"/>
                  <TextBlock Grid.Row="0" Grid.Column="1" x:Name="LblStatus" Text="Not Connected" Margin="0,0,0,4" Foreground="#111827" FontWeight="Bold"/>

                  <TextBlock Grid.Row="1" Grid.Column="0" Text="User:" Style="{StaticResource SidebarText}" FontWeight="SemiBold" Foreground="#111827" Margin="0,0,6,4"/>
                  <TextBlock Grid.Row="1" Grid.Column="1" x:Name="LblUser" Text="-" Margin="0,0,0,4" Foreground="#4B5563" TextWrapping="Wrap"/>

                  <TextBlock Grid.Row="2" Grid.Column="0" Text="TenantId:" Style="{StaticResource SidebarText}" FontWeight="SemiBold" Foreground="#111827"/>
                  <TextBlock Grid.Row="2" Grid.Column="1" x:Name="LblTenant" Text="-" Foreground="#4B5563" TextWrapping="Wrap"/>
                </Grid>

                <Grid Margin="0,10,0,0">
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Button Grid.Column="0" x:Name="BtnGraphConnect" Content="Connect" Style="{StaticResource BtnPrimary}"/>
                  <Button Grid.Column="2" x:Name="BtnGraphDisconnect" Content="Disconnect" Style="{StaticResource BtnNeutral}"/>
                </Grid>
              </StackPanel>
            </Border>

            <!-- Session -->
            <Border Style="{StaticResource SidebarCard}" Margin="12,0,12,8">
              <StackPanel>
                <TextBlock Text="Session" Style="{StaticResource SidebarTitle}"/>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>

                  <TextBlock Grid.Row="0" Grid.Column="0" Text="Machine:" Style="{StaticResource SidebarText}" FontWeight="SemiBold" Foreground="#111827" Margin="0,0,6,4"/>
                  <Border Grid.Row="0" Grid.Column="1" Background="#EEF2FF" Padding="6,2" CornerRadius="4" Margin="0,0,0,4">
                    <TextBlock x:Name="LblSessionMachine" Text="-" Foreground="#1D4ED8"/>
                  </Border>

                  <TextBlock Grid.Row="1" Grid.Column="0" Text="User:" Style="{StaticResource SidebarText}" FontWeight="SemiBold" Foreground="#111827" Margin="0,0,6,4"/>
                  <Border Grid.Row="1" Grid.Column="1" Background="#ECFDF3" Padding="6,2" CornerRadius="4" Margin="0,0,0,4">
                    <TextBlock x:Name="LblSessionUser" Text="-" Foreground="#166534"/>
                  </Border>

                  <TextBlock Grid.Row="2" Grid.Column="0" Text="Elevation:" Style="{StaticResource SidebarText}" FontWeight="SemiBold" Foreground="#111827" Margin="0,0,6,0"/>
                  <Border Grid.Row="2" Grid.Column="1" Background="#ECFDF3" Padding="6,2" CornerRadius="4">
                    <TextBlock x:Name="LblSessionElevation" Text="-" Foreground="#166534"/>
                  </Border>
                </Grid>
              </StackPanel>
            </Border>

            <!-- About -->
            <Border Style="{StaticResource SidebarCard}" Margin="12,0,12,12">
              <StackPanel>
                <TextBlock Text="About this assistant" FontSize="12" FontWeight="SemiBold" Foreground="#1F2D3A" Margin="0,0,0,6"/>
                <TextBlock Foreground="#475467" FontSize="12" TextWrapping="Wrap"
                           Text="Collect HWID (CSV), connect to Microsoft Graph, and import devices to Windows Autopilot from one console."/>
              </StackPanel>
            </Border>
          </StackPanel>
        </Grid>

      </DockPanel>
    </Border>

    <!-- Main Content -->
    <Grid Grid.Column="1">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>

      <!-- Header -->
      <Border Grid.Row="0" Padding="18,14,18,10" Background="#F6F8FB">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0">
            <TextBlock Text="Welcome" FontSize="20" FontWeight="Bold" Foreground="#1F2D3A"/>
            <TextBlock Text="Collect HWID, connect to Graph, then upload/import to Autopilot"
                       FontSize="14" Foreground="#5F6B7A" Margin="0,6,0,0"/>
          </StackPanel>
          <Button Grid.Column="1" x:Name="BtnRefreshDevice" Content="Refresh Device Info"
                  Height="32" MinWidth="160" Margin="12,0,0,0"
                  Style="{StaticResource BtnPrimary}"/>
        </Grid>
      </Border>

      <!-- Body -->
      <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
        <Grid Margin="16,0,16,16">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <!-- Device Information -->
          <Border Grid.Row="0" Background="White" CornerRadius="5" BorderBrush="#E4E9F0" BorderThickness="1" Padding="12" Margin="0,0,0,5">
            <StackPanel>
              <TextBlock Text="Device Information" FontSize="13" FontWeight="SemiBold" Foreground="#0F172A" Margin="0,0,0,10"/>

              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="150"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="130"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="5"/>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="5"/>
                  <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Grid.Column="0" Text="Device Model:" Foreground="#1F2D3A" FontWeight="SemiBold"/>
                <Border Grid.Row="0" Grid.Column="1" Style="{StaticResource DeviceValuePill}">
                  <TextBlock x:Name="LblDevModel" Text="-" Style="{StaticResource DeviceValueText}"/>
                </Border>

                <TextBlock Grid.Row="0" Grid.Column="2" Text="Free Storage (GB):" Foreground="#1F2D3A" FontWeight="SemiBold"/>
                <Border Grid.Row="0" Grid.Column="3" Style="{StaticResource DeviceValuePill}">
                  <TextBlock x:Name="LblFreeGb" Text="-" Style="{StaticResource DeviceValueText}"/>
                </Border>

                <TextBlock Grid.Row="2" Grid.Column="0" Text="Device Name:" Foreground="#1F2D3A" FontWeight="SemiBold"/>
                <Border Grid.Row="2" Grid.Column="1" Style="{StaticResource DeviceValuePill}">
                  <TextBlock x:Name="LblDevName" Text="-" Style="{StaticResource DeviceValueText}"/>
                </Border>

                <TextBlock Grid.Row="2" Grid.Column="2" Text="TPM Version:" Foreground="#1F2D3A" FontWeight="SemiBold"/>
                <Border Grid.Row="2" Grid.Column="3" Style="{StaticResource DeviceValuePill}">
                  <TextBlock x:Name="LblTpm" Text="-" Style="{StaticResource DeviceValueText}"/>
                </Border>

                <TextBlock Grid.Row="4" Grid.Column="0" Text="Manufacturer:" Foreground="#1F2D3A" FontWeight="SemiBold"/>
                <Border Grid.Row="4" Grid.Column="1" Style="{StaticResource DeviceValuePill}">
                  <TextBlock x:Name="LblManufacturer" Text="-" Style="{StaticResource DeviceValueText}"/>
                </Border>

                <TextBlock Grid.Row="4" Grid.Column="2" Text="Internet:" Foreground="#1F2D3A" FontWeight="SemiBold"/>
                <Border Grid.Row="4" Grid.Column="3" Style="{StaticResource DeviceValuePill}">
                  <TextBlock x:Name="LblNet" Text="-" Style="{StaticResource DeviceValueText}"/>
                </Border>
              </Grid>

              <Grid Margin="0,5,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="150"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="Serial Number:" Foreground="#1F2D3A" FontWeight="SemiBold"/>
                <Border Grid.Column="1" Style="{StaticResource DeviceValuePill}">
                  <TextBlock x:Name="LblSerial" Text="-" Style="{StaticResource DeviceValueText}"/>
                </Border>
              </Grid>

            </StackPanel>
          </Border>

          <!-- Autopilot Actions -->
          <Grid Grid.Row="1" Margin="0,3,0,8">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="340"/>
              <ColumnDefinition Width="10"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Method 1: Local CSV export only -->
            <Border Grid.Column="0" Background="White" CornerRadius="5" BorderBrush="#E4E9F0" BorderThickness="1" Padding="12">
              <StackPanel>
                <TextBlock Text="Method 1: Collect HWID to CSV (No Upload)" FontSize="13" FontWeight="SemiBold" Foreground="#0F172A"/>
                <TextBlock Text="Use this method when you only need a CSV export. It collects serial/product/hash locally and does not upload to Intune."
                           Foreground="#5F6B7A" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,10"/>

                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>

                  <TextBlock Grid.Row="0" Text="Save folder:" VerticalAlignment="Center" Foreground="#1F2D3A" Margin="0,0,0,6"/>

                  <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <TextBox Grid.Column="0" x:Name="TxtSaveFolder" MinWidth="200" Margin="0,0,8,0"/>
                    <Button  Grid.Column="1" x:Name="BtnBrowseFolder" Content="Browse..." Width="96" Style="{StaticResource BtnSmall}"/>
                  </Grid>

                  <Button Grid.Row="2" x:Name="BtnCollectHash" Margin="0,8,0,0"
                          ToolTip="Collect local hardware hash and save a CSV in the selected folder."
                          Content="Collect Hash" Style="{StaticResource BtnPrimary}"/>
                </Grid>
              </StackPanel>
            </Border>

            <!-- Method 2: Direct upload/import -->
            <Border Grid.Column="2" Background="White" CornerRadius="5" BorderBrush="#E4E9F0" BorderThickness="1" Padding="12">
              <StackPanel>
                <TextBlock Text="Method 2: Upload/Import Device to Intune Autopilot" FontSize="13" FontWeight="SemiBold" Foreground="#0F172A"/>
                <TextBlock Text="Upload device records directly to Intune Autopilot. If CSV path is empty, missing, or a folder path, the app auto-creates a CSV in the default HWID folder."
                           Foreground="#5F6B7A" FontSize="11" TextWrapping="Wrap" Margin="0,2,0,10"/>

                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="160"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="10"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                  </Grid.RowDefinitions>

                  <TextBlock Grid.Row="0" Grid.Column="0" Text="CSV file (optional):" VerticalAlignment="Center" Foreground="#1F2D3A"/>
                  <TextBox   Grid.Row="0" Grid.Column="1" x:Name="TxtImportPath"/>
                  <Button    Grid.Row="0" Grid.Column="2" x:Name="BtnBrowseImport" Content="Browse..." Margin="8,0,0,0" Style="{StaticResource BtnSmall}"/>

                  <TextBlock Grid.Row="2" Grid.Column="0" Text="Group Tag:" VerticalAlignment="Center" Foreground="#1F2D3A"/>
                  <TextBox   Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" x:Name="TxtGroupTag"/>

                  <TextBlock Grid.Row="4" Grid.Column="0" Text="Assigned User:" VerticalAlignment="Center" Foreground="#1F2D3A"/>
                  <TextBox   Grid.Row="4" Grid.Column="1" Grid.ColumnSpan="2" x:Name="TxtAssignedUser"/>

                  <TextBlock Grid.Row="6" Grid.Column="0" Text="Assigned Computer Name:" VerticalAlignment="Center" Foreground="#1F2D3A"/>
                  <TextBox   Grid.Row="6" Grid.Column="1" Grid.ColumnSpan="2" x:Name="TxtAssignedName"/>

                  <Grid Grid.Row="8" Grid.Column="0" Grid.ColumnSpan="3">
                    <Grid.RowDefinitions>
                      <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Button Grid.Row="0" Grid.Column="0" x:Name="BtnEnroll" Content="Upload Device to Autopilot" Margin="0,0,8,0"
                            ToolTip="Upload current/imported records to Intune Autopilot. Auto-creates CSV when input path is empty/missing/folder."
                            Style="{StaticResource BtnPrimary}"/>
                    <Button Grid.Row="0" Grid.Column="1" x:Name="BtnRetryFailed"
                            Content="Retry Failed Uploads"
                            IsEnabled="False"
                            ToolTip="Retry only failed rows from the most recent upload attempt."
                            Style="{StaticResource BtnNeutral}"/>
                  </Grid>

                  <TextBlock Grid.Row="10" Grid.Column="0" Text="Upload progress:" VerticalAlignment="Center" Foreground="#1F2D3A"/>
                  <ProgressBar x:Name="ProgressBar"
                               Grid.Row="10"
                               Grid.Column="1"
                               Grid.ColumnSpan="2"
                               Value="0"
                               Maximum="100"
                               Height="12"
                               VerticalAlignment="Center"
                               Visibility="Collapsed"/>
                </Grid>

              </StackPanel>
            </Border>
          </Grid>

          <!-- Message Center -->
          <Border Grid.Row="2" Background="White" CornerRadius="5" BorderBrush="#E4E9F0" BorderThickness="1" Padding="12">
            <StackPanel>
              <Grid Margin="0,0,0,8">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Grid.Column="0" Text="MESSAGE CENTER" FontSize="13" FontWeight="SemiBold" Foreground="#0F172A"/>
                <Button Grid.Column="1" x:Name="BtnClearLog" Content="Clear" Margin="8,0,0,0" Style="{StaticResource BtnSmall}"/>
                <Button Grid.Column="2" x:Name="BtnCopyLog"  Content="Copy"  Margin="8,0,0,0" Style="{StaticResource BtnSmall}"/>
              </Grid>

              <RichTextBox x:Name="TxtLog"
                           Height="140"
                           IsReadOnly="True"
                           Background="#1F2D3A"
                           Foreground="#E4E9F0"
                           BorderBrush="#1F2937"
                           BorderThickness="1"
                           FontFamily="Consolas"
                           FontSize="13"
                           VerticalScrollBarVisibility="Auto"
                           Padding="10"/>
            </StackPanel>
          </Border>

        </Grid>
      </ScrollViewer>
    </Grid>
  </Grid>
</Window>
"@

# Build main window and capture controls
$Window          = [Windows.Markup.XamlReader]::Parse($Xaml)

$FooterLeft      = $Window.FindName('FooterLeft')
$FooterCenter    = $Window.FindName('FooterCenter')
$FooterCopyRun   = $Window.FindName('FooterCopyRun')
$FooterLink      = $Window.FindName('FooterLink')


# Device info
$LblDevModel     = $Window.FindName('LblDevModel')
$LblDevName      = $Window.FindName('LblDevName')
$LblManufacturer = $Window.FindName('LblManufacturer')
$LblSerial       = $Window.FindName('LblSerial')
$LblFreeGb       = $Window.FindName('LblFreeGb')
$LblTpm          = $Window.FindName('LblTpm')
$LblNet          = $Window.FindName('LblNet')

# Graph
$BtnGraphConnect    = $Window.FindName('BtnGraphConnect')
$BtnGraphDisconnect = $Window.FindName('BtnGraphDisconnect')
$LblStatus          = $Window.FindName('LblStatus')
$LblUser            = $Window.FindName('LblUser')
$LblTenant          = $Window.FindName('LblTenant')
$SideGraphPill      = $Window.FindName('SideGraphPill')
$SideGraphTxt       = $Window.FindName('SideGraphTxt')

# Header refresh
$BtnRefreshDevice   = $Window.FindName('BtnRefreshDevice')

# Sidebar session labels
$LblSessionMachine  = $Window.FindName('LblSessionMachine')
$LblSessionUser     = $Window.FindName('LblSessionUser')
$LblSessionElevation= $Window.FindName('LblSessionElevation')

# Autopilot controls
$TxtSaveFolder      = $Window.FindName('TxtSaveFolder')
$BtnBrowseFolder    = $Window.FindName('BtnBrowseFolder')
$BtnCollectHash     = $Window.FindName('BtnCollectHash')

$TxtImportPath      = $Window.FindName('TxtImportPath')
$BtnBrowseImport    = $Window.FindName('BtnBrowseImport')
$TxtGroupTag        = $Window.FindName('TxtGroupTag')
$TxtAssignedUser    = $Window.FindName('TxtAssignedUser')
$TxtAssignedName    = $Window.FindName('TxtAssignedName')
$BtnEnroll          = $Window.FindName('BtnEnroll')
$BtnRetryFailed     = $Window.FindName('BtnRetryFailed')
$ProgressBar        = $Window.FindName('ProgressBar')

# Message Center
$BtnClearLog        = $Window.FindName('BtnClearLog')
$BtnCopyLog         = $Window.FindName('BtnCopyLog')
$TxtLog             = $Window.FindName('TxtLog'); $TxtLog.Document = New-Object Windows.Documents.FlowDocument

# Footer
$FooterLeft.Text   = "Autopilot Assistant"
$FooterCenter.Text = "Version $AppVersion"
if ($FooterCopyRun) { $FooterCopyRun.Text = "(c) 2025 " }
if ($FooterLink) {
  $FooterLink.Add_RequestNavigate({
    param($sender, $e)
    try {
      $psi = New-Object System.Diagnostics.ProcessStartInfo
      $psi.FileName = $e.Uri.AbsoluteUri
      $psi.UseShellExecute = $true
      [System.Diagnostics.Process]::Start($psi) | Out-Null
    } catch { }
    $e.Handled = $true
  })
}

#endregion ============================== MAIN WINDOW UI ==============================================

#region ============================== WIRE UP UI EVENTS =============================================
$Window.Add_Loaded({
  Refresh-DeviceInfo
  Refresh-SessionInfo
  Update-MainGraphUI
  $TxtSaveFolder.Text = $Paths.HwId
  $TxtImportPath.Text = $Paths.HwId
  $script:LastFailedRecords = @()
  Stop-UploadProgress
  if ($BtnRetryFailed) { $BtnRetryFailed.IsEnabled = $false }
  Add-Log "Ready." "INFO"
  Add-Log "Method 1: collect HWID to CSV only (no Graph upload)." "INFO"
  Add-Log "Method 2: upload/import to Intune Autopilot; CSV is auto-created when input path is empty, missing, or folder." "INFO"
  if (-not $BtnGraphConnect -or -not $BtnGraphDisconnect) {
    Add-Log "Graph buttons were not initialized from XAML (BtnGraphConnect/BtnGraphDisconnect)." "ERROR"
  }
})

$Window.Add_Closing({
  Invoke-AppCleanup
})

$Window.Add_Closed({
  Invoke-AppCleanup
})

$BtnClearLog.Add_Click({ $TxtLog.Document.Blocks.Clear() })
$BtnCopyLog.Add_Click({
  $range = New-Object Windows.Documents.TextRange($TxtLog.Document.ContentStart, $TxtLog.Document.ContentEnd)
  [System.Windows.Clipboard]::SetText($range.Text)
})

$BtnRefreshDevice.Add_Click({
  Refresh-DeviceInfo
  Refresh-SessionInfo
  Add-Log "Device info refreshed." "INFO"
})

$BtnBrowseFolder.Add_Click({
  if (-not (Test-Path -LiteralPath $Paths.HwId)) {
    New-Item -ItemType Directory -Path $Paths.HwId -Force | Out-Null
  }
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.Description = 'Select a folder to save the Autopilot CSV'
  $dlg.SelectedPath = $Paths.HwId
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $TxtSaveFolder.Text = $dlg.SelectedPath }
})

$BtnCollectHash.Add_Click({
  try {
    if (-not (Run-PreChecks -Mode 'collect')) { return }
    Add-Log "Collecting hardware hash..." "INFO"
    $BtnCollectHash.IsEnabled = $false
    $script:CurrentAction     = 'collect'

    $script:PS = [System.Management.Automation.PowerShell]::Create()
    $script:PS.RunspacePool = $pool
    $script:PS.AddScript($CollectHWIDWorker).AddArgument($TxtSaveFolder.Text) | Out-Null
    $script:Async = $script:PS.BeginInvoke()
    $Timer.Start()
  } catch {
    Add-Log ("Collect start error: " + $_.Exception.Message) "ERROR"
    try { $script:PS.Dispose() } catch { }
    $BtnCollectHash.IsEnabled = $true
  }
})

$BtnBrowseImport.Add_Click({
  if (-not (Test-Path -LiteralPath $Paths.HwId)) {
    New-Item -ItemType Directory -Path $Paths.HwId -Force | Out-Null
  }
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Title  = 'Select Autopilot CSV or JSON'
  $dlg.Filter = 'CSV or JSON|*.csv;*.json|CSV|*.csv|JSON|*.json|All files|*.*'
  $startDir = $Paths.HwId
  if (-not [string]::IsNullOrWhiteSpace($TxtImportPath.Text)) {
    if (Test-Path -LiteralPath $TxtImportPath.Text -PathType Container) {
      $startDir = $TxtImportPath.Text
    } elseif (Test-Path -LiteralPath $TxtImportPath.Text -PathType Leaf) {
      try {
        $startDir = Split-Path -Path $TxtImportPath.Text -Parent
      } catch { }
    }
  }
  $dlg.InitialDirectory = $startDir
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $TxtImportPath.Text = $dlg.FileName }
})

$BtnEnroll.Add_Click({
  try {
    Add-Log "Starting upload workflow: validating Graph module and connection..." "INFO"
    if (-not (Run-PreChecks -Mode 'enroll')) { return }

    $candidates = Get-UploadSerialCandidates -Path $TxtImportPath.Text -FallbackSerial $LblSerial.Text
    if ($candidates.Count -gt 0) {
      Add-Log ("Pre-check: verifying existing Autopilot records for " + $candidates.Count + " serial...") "INFO"
      $hardExists = @()
      foreach ($s in $candidates) {
        $chk = Test-AutopilotSerialExists -Serial $s
        if ($chk.Exists) { $hardExists += $chk }
      }
      if ($hardExists.Count -gt 0) {
        $existingSerials = @()
        foreach ($e in $hardExists) {
          Add-Log ("Pre-check: serial already exists in " + $e.Source + " -> " + $e.Serial) "WARN"
          if (-not [string]::IsNullOrWhiteSpace([string]$e.Serial)) { $existingSerials += [string]$e.Serial }
        }
        # Local one-device workflow: if device already exists, do not upload again.
        Add-Log "Upload skipped: this device is already registered in Autopilot." "WARN"
        Show-UploadBlockedDeviceExists -Serials @($existingSerials | Select-Object -Unique)
        return
      }
    } else {
      Add-Log "Pre-check warning: no serial candidates found before upload. Continuing..." "WARN"
    }

    Add-Log "Uploading to Autopilot..." "INFO"
    $BtnEnroll.IsEnabled  = $false
    if ($BtnRetryFailed) { $BtnRetryFailed.IsEnabled = $false }
    $script:CurrentAction = 'enroll'
    Start-UploadProgress

    $script:PS = [System.Management.Automation.PowerShell]::Create()
    $script:PS.RunspacePool = $pool
    $script:PS.AddScript($EnrollWorker).
      AddArgument($TxtImportPath.Text).
      AddArgument($TxtGroupTag.Text).
      AddArgument($TxtAssignedUser.Text).
      AddArgument($TxtAssignedName.Text).
      AddArgument($Paths.HwId).
      AddArgument($null) | Out-Null

    $script:Async = $script:PS.BeginInvoke()
    $Timer.Start()
  } catch {
    Add-Log ("Upload start error: " + $_.Exception.Message) "ERROR"
    Stop-UploadProgress
    try { $script:PS.Dispose() } catch { }
    $BtnEnroll.IsEnabled = $true
  }
})

$BtnRetryFailed.Add_Click({
  try {
    if (-not $script:LastFailedRecords -or $script:LastFailedRecords.Count -eq 0) {
      Add-Log "No failed rows to retry." "WARN"
      return
    }
    Add-Log "Starting retry workflow: validating Graph module and connection..." "INFO"
    if (-not (Run-PreChecks -Mode 'enroll')) { return }

    $retryJson = $script:LastFailedRecords | ConvertTo-Json -Depth 8 -Compress
    Add-Log ("Retrying failed rows: " + $script:LastFailedRecords.Count) "INFO"

    $BtnEnroll.IsEnabled = $false
    $BtnRetryFailed.IsEnabled = $false
    $script:CurrentAction = 'retry'
    Start-UploadProgress

    $script:PS = [System.Management.Automation.PowerShell]::Create()
    $script:PS.RunspacePool = $pool
    $script:PS.AddScript($EnrollWorker).
      AddArgument($null).
      AddArgument($TxtGroupTag.Text).
      AddArgument($TxtAssignedUser.Text).
      AddArgument($TxtAssignedName.Text).
      AddArgument($Paths.HwId).
      AddArgument($retryJson) | Out-Null

    $script:Async = $script:PS.BeginInvoke()
    $Timer.Start()
  } catch {
    Add-Log ("Retry start error: " + $_.Exception.Message) "ERROR"
    Stop-UploadProgress
    try { $script:PS.Dispose() } catch { }
    $BtnEnroll.IsEnabled = $true
    if ($BtnRetryFailed) { $BtnRetryFailed.IsEnabled = ($script:LastFailedRecords.Count -gt 0) }
  }
})

if ($BtnGraphConnect) {
  $BtnGraphConnect.Add_Click({
    Add-Log "Connect button clicked." "INFO"
    try {
      Connect-GraphBrowserInteractive
    } catch {
      Add-Log ("Graph connect handler error: " + $_.Exception.Message) "ERROR"
    }
  })
}

if ($BtnGraphDisconnect) {
  $BtnGraphDisconnect.Add_Click({
    Add-Log "Disconnect button clicked." "INFO"
    try {
      Disconnect-GraphDirect
    } catch {
      Add-Log ("Graph disconnect handler error: " + $_.Exception.Message) "ERROR"
    }
  })
}

#endregion ============================== WIRE UP UI EVENTS ===========================================

#region ============================== RUN THE UI ====================================================
[void]$Window.ShowDialog()
Invoke-AppCleanup
try {
  if ($PSCommandPath -and $PSCommandPath.ToLower().EndsWith('.exe')) {
    [System.Environment]::Exit(0)
  }
} catch { }
#endregion ============================== RUN THE UI ==================================================




