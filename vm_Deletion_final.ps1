#requires -Version 5.1

### Created By  Tene Bidner ###
### January 2026 ###
### Version:1.0
<#
.SYNOPSIS
    GUI tool to delete a VM/Template and clean up related records from vCenter, SCCM, AD, and DNS.

.DESCRIPTION
    VM Deletion Tool supports two operating modes:

    1. Interactive GUI mode
       Run the script without parameters.
       The GUI lets the operator select a vCenter, load VMs, import a CSV list,
       select deletion actions, confirm deletion, review progress, and check archive data.

       Example:
       powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Tene\Codex\Scirpts_Tene\Vm_deletion\vm_Deletion_v1.0.ps1"

    2. Scheduled / headless mode for Task Scheduler
       Run the script with -NoGui and -CsvPath.
       No GUI windows are opened.
       The script reads a CSV file, groups servers by Environment/vCenter,
       connects to the relevant vCenter when VM deletion is enabled,
       performs the requested cleanup steps, writes logs, and returns an exit code.

       Example:
       powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Tene\Codex\Scirpts_Tene\Vm_deletion\vm_Deletion_v1.0.ps1" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD

       Safe validation example without deleting anything:
       powershell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Tene\Codex\Scirpts_Tene\Vm_deletion\vm_Deletion_v1.0.ps1" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD -SkipSCCM -SkipAD -SkipDNS -SkipVM -SkipArchive -SkipEmail

    Exit codes in NoGui mode:
       0 = completed without VM-level failures
       1 = one or more failures occurred, or startup validation failed

    Required CSV structure:
       VMName
       SERVER01
       SERVER02

    Recommended CSV structure:
       VMName,Environment
       SERVER01,PROD
       SERVER02,MGMT
       SERVER03,VDI

    Alternative CSV structure:
       VMName,vCenter
       SERVER01,Prod
       SERVER02,Mgmt
       SERVER03,Vdi

    CSV columns:
       VMName      Required. VM or Template name to process.
       Environment Optional. Used to select vCenter and AD OU policy.
       vCenter     Optional. Can be used instead of Environment.

    Environment / AD OU behavior:
       AD deletion uses ADSearchBaseByEnvironment when available.
       PROD and MGMT use the same default OU.
       VDI uses a separate default OU.
       These values can be overridden in config.json.

    Default AD OU map in code:
       ADSearchBaseByEnvironment = @{
           PROD = "OU=Servers,DC=system,DC=co,DC=il"
           MGMT = "OU=Servers,DC=system,DC=co,DC=il"
           VDI  = "OU=VDI,DC=system,DC=co,DC=il"
       }

    Configuration files:
       Local JSON:
           $env:TEMP\VMDeletionTool\config.json

       Network JSON:
           Tene\VMDeletionTool\config.json

       Startup behavior:
           The script waits briefly for JSON availability.
           If no local or network JSON is found, it uses the default settings embedded in the code.
           In GUI mode, a short auto-closing notice is shown.
           In NoGui mode, no notice window is shown.

    Main cleanup actions:
       SCCM    Delete SCCM device record.
       AD      Delete computer object from the OU selected by Environment.
       DNS     Delete A/PTR/AAAA records.
       VM      Delete VM or Template from vCenter.
       Archive Write result to SQL archive table when SQL is configured and checked.
       Email   Send deletion notification when email is enabled.

    NoGui skip switches:
       -SkipSCCM
       -SkipAD
       -SkipDNS
       -SkipVM
       -SkipArchive
       -SkipEmail

    Task Scheduler notes:
       Use a service account with the required permissions for the enabled actions.
       VMware PowerCLI is required when VM deletion is not skipped.
       ActiveDirectory module is required when AD deletion is not skipped.
       DnsServer module is required when DNS deletion is not skipped.
       SCCM access is required when SCCM deletion is not skipped.
       SQL connectivity is required when Archive is not skipped.

.PARAMETER Help
    Shows NoGui usage examples, CSV format, and exit codes, then exits without opening the GUI.

.PARAMETER NoGui
    Runs in headless mode without opening the GUI. Intended for Task Scheduler.

.PARAMETER CsvPath
    Path to a CSV file containing at least a VMName column.

.PARAMETER Environment
    Default environment used when the CSV does not include Environment or vCenter.
    Common values: PROD, MGMT, VDI.

.PARAMETER Reason
    Reason text written to archive/email in NoGui mode.

.PARAMETER SkipSCCM
    Skips SCCM cleanup.

.PARAMETER SkipAD
    Skips AD cleanup.

.PARAMETER SkipDNS
    Skips DNS cleanup.

.PARAMETER SkipVM
    Skips VM/Template deletion from vCenter.

.PARAMETER SkipArchive
    Skips SQL archive write.

.PARAMETER SkipEmail
    Skips email notification.

.NOTES
    Requires: VMware PowerCLI, SCCM module, ActiveDirectory module, DNS Server tools
    v2.1 - Added Template deletion support + Linux cluster email logic
#>
param(
    [switch]$Help,
    [switch]$NoGui,
    [string]$CsvPath,
    [string]$Environment = "PROD",
    [string]$Reason = "Scheduled CSV deletion",
    [switch]$SkipSCCM,
    [switch]$SkipAD,
    [switch]$SkipDNS,
    [switch]$SkipVM,
    [switch]$SkipArchive,
    [switch]$SkipEmail
)

CLS

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

#region Global Variables and Settings

# ============================================
# Configuration Sources
# ============================================
# 1. Shared JSON config - primary source of truth.
$script:SharedConfigPath = '\\teneserver\system\shared\scripts\vmdeletion\tools\config.json'
$script:SharedConfigTimeoutMs = 8000

# 2. Per-user TEMP JSON cache - used when shared config is unavailable.
$script:LocalConfigCacheDir = Join-Path $env:TEMP 'VMDeletionTool'
$script:LocalConfigCachePath = Join-Path $script:LocalConfigCacheDir 'config.json'
$script:LocalConfigTimeoutMs = 1000

# 3. Built-in defaults are used only if shared and local JSON are unavailable or invalid.
$script:ConfigSource = ''
# ============================================
# Global Variables
# ============================================
$script:vcConnected = $false
$script:allVMs = @()
$script:sqlTableExists = $false
$script:vCenterServers = @()
$script:DefaultSettings = @{}
$script:SelectedVMNames = [System.Collections.ArrayList]@()
$script:IsUpdatingList = $false
$script:allVMsData = @()
$script:allTemplatesData = @()
$script:IncludeTemplates = $false
$script:settingsPassword = "admin"
$script:RunEnvironment = $Environment
$script:CSVImportData = @()
$script:txtLogs = $null
$script:logsTab = $null
$script:IsShowingLogErrorNotice = $false

$tasks = @("Connect to vCenter", "Delete SCCM Record", "Delete AD Record", "Delete DNS Record", "Delete VM", "Archive to SQL", "Send Notification")
$statusLabels = @()
$resultLabels = @()
$yPos = 50

# ============================================
# Configuration Functions
# ============================================

function CheckPermission {
    $memebers = ((Get-ADGroup "System" -ErrorAction SilentlyContinue | Get-ADGroupMember -ErrorAction SilentlyContinue).SamAccountName)
    if ($memebers.count -gt 0 -and $env:USERNAME -match ($memebers -join "|")){
     return $True
     }else{
        return $False
    }
}

function Wait-JsonFileAvailable {
    param(
        [string]$Path,
        [int]$TimeoutMs = 1000,
        [int]$IntervalMs = 100
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    do {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            return $Path
        }

        Start-Sleep -Milliseconds $IntervalMs
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)

    return $null
}

function Show-AutoCloseNotice {
    param(
        [string]$Message,
        [string]$Title = "VM Deletion Tool",
        [int]$DurationMs = 2200
    )

    try {
        $noticeForm = New-Object System.Windows.Forms.Form
        $noticeForm.Text = $Title
        $noticeForm.Size = New-Object System.Drawing.Size(460, 120)
        $noticeForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $noticeForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedToolWindow
        $noticeForm.MaximizeBox = $false
        $noticeForm.MinimizeBox = $false
        $noticeForm.TopMost = $true
        $noticeForm.BackColor = [System.Drawing.Color]::FromArgb(54, 54, 68)

        $noticeLabel = New-Object System.Windows.Forms.Label
        $noticeLabel.Location = New-Object System.Drawing.Point(18, 18)
        $noticeLabel.Size = New-Object System.Drawing.Size(410, 48)
        $noticeLabel.Text = $Message
        $noticeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $noticeLabel.ForeColor = [System.Drawing.Color]::FromArgb(245, 247, 250)
        $noticeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $noticeForm.Controls.Add($noticeLabel)

        $noticeTimer = New-Object System.Windows.Forms.Timer
        $noticeTimer.Interval = $DurationMs
        $noticeTimer.Add_Tick({
            $noticeTimer.Stop()
            $noticeForm.Close()
        })
        $noticeForm.Add_Shown({ $noticeTimer.Start() })
        [void]$noticeForm.ShowDialog()
        $noticeTimer.Dispose()
        $noticeForm.Dispose()
    }
    catch {
        return
    }
}

function Load-ConfigFromFile {
    param(
        [string]$Path,
        [int]$TimeoutMs = 1000
    )
    
    try {
        $availablePath = Wait-JsonFileAvailable -Path $Path -TimeoutMs $TimeoutMs
        if (-not $availablePath) { return $null }
        
        $jsonContent = Get-Content -LiteralPath $availablePath -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($jsonContent)) { return $null }

        return ($jsonContent | ConvertFrom-Json)
    }
    catch {
        ##Write-Host "Failed to load config from $Path : $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Apply-Config {
    param($Config)

    try {
        if ($null -eq $Config) { return $false }

        $appliedConfig = $false
        $configPropertyNames = @($Config.PSObject.Properties.Name)
        $settingsSource = if ($configPropertyNames -contains "DefaultSettings") {
            $Config.DefaultSettings
        }
        else {
            $Config
        }

        if ($configPropertyNames -contains "vCenterServers" -and $Config.vCenterServers) {
            $loadedVCenterServers = @()
            foreach ($vc in $Config.vCenterServers) {
                if ([string]::IsNullOrWhiteSpace($vc.Name) -or [string]::IsNullOrWhiteSpace($vc.Server)) { continue }
                $loadedVCenterServers += @{ Name = $vc.Name; Server = $vc.Server }
            }

            if ($loadedVCenterServers.Count -gt 0) {
                $script:vCenterServers = $loadedVCenterServers
                $appliedConfig = $true
            }
        }

        if ($settingsSource) {
            foreach ($property in $settingsSource.PSObject.Properties) {
                if ($property.Name -in @("vCenterServers", "DefaultSettings", "LastSaved", "SavedBy")) { continue }
                $script:DefaultSettings[$property.Name] = $property.Value
                $appliedConfig = $true
            }
        }

        if (-not $script:DefaultSettings -or $script:DefaultSettings.Count -eq 0) { return $false }

        if ([string]::IsNullOrWhiteSpace($script:DefaultSettings.LogPath)) {
            $script:DefaultSettings.LogPath = $script:LocalConfigCacheDir
        }

        return $appliedConfig
    }
    catch {
        ##Write-Host "Failed to apply config: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}
function Initialize-Configuration {
    $script:ConfigSource = ""
    Set-DefaultConfiguration

    # 1. Try the shared network config first.
    $networkConfig = Load-ConfigFromFile -Path $script:SharedConfigPath -TimeoutMs $script:SharedConfigTimeoutMs
    if ($networkConfig -and (Apply-Config -Config $networkConfig)) {
        $script:ConfigSource = "Shared JSON: $script:SharedConfigPath"
        Save-LocalConfig | Out-Null
        return $true
    }

    # 2. If shared config is unavailable, use the local TEMP config.
    $localConfig = Load-ConfigFromFile -Path $script:LocalConfigCachePath -TimeoutMs $script:LocalConfigTimeoutMs
    if ($localConfig -and (Apply-Config -Config $localConfig)) {
        $script:ConfigSource = "TEMP JSON: $script:LocalConfigCachePath"
        return $true
    }

    # 3. If local TEMP is unavailable or invalid, keep the in-memory built-in defaults.
    $script:ConfigSource = "Built-in defaults from script"
    if (-not $NoGui) {
        Show-AutoCloseNotice -Message "Network and TEMP JSON config are unavailable.`nUsing default settings from code."
    }
    return $true
}
function Set-DefaultConfiguration {
    # רשימת שרתי vCenter
    $script:vCenterServers = @(
    @{ Name = "Prod";  Server = "VC1" },
    @{ Name = "Mgmt"; Server = "VC2" },
    @{ Name = "Vdi";        Server = "VC3" }
)

    $script:DefaultSettings = @{
    vCenterServer           = "vcenter01.domain.local"
    SCCMServer              = "sccmsrv.co.il"
    SCCMSiteCode            = "ABC"
    DNSServer               = "DNS1"
    DNSZone                 = "dns.co.il"
    DomainController        = "DC1"
    EmailNotification       = "windows.co.il" 
    LogPath                 = "$env:TEMP\VMDeletionTool"
    EnableLogging           = $true
    ConfirmDeletion         = $true
    PowerOffTimeout         = 30
    DeleteFromDisk          = $true
    

    # SQL Settings
    SQLServer               = "YOUR_SQL_SERVER"
    SQLDatabase             = "YOUR_SQL_DATABASE"
    SQLTable                = "YOUR_SQL_TABLE"
    SQLIntegratedAuth       = $false
    SQLUsername             = ""YOUR_SQL_USERNAME""
    SQLPassword             = ""YOUR_SQL_PASSWORD""

    # Email Settings
    SMTPServer              = "YOUR_SMTP"
    SMTPPort                = 25
    SMTPUseSSL              = $false
    SMTPFrom                = "Sender"
    EmailSubjectPrefix      = "[VM Deletion]"
    EnableEmailNotification = $true
    
    #AD-Delete by OU
    ADSearchBase = "OU=Servers,DC=system,DC=co,DC=il"
    ADSearchBaseByEnvironment = @{
        PROD = "OU=Servers,DC=system,DC=co,DC=il"
        MGMT = "OU=Servers,DC=system,DC=co,DC=il"
        VDI  = "OU=Servers,OU=VDI,DC=system,DC=co,DC=il"
    }
   
    # AD Groups for Email (comma-separated)
    GroupsCommon            = "test@system.co.il"
    GroupsWindows           = "test@system.co.il" 
    GroupsLinux             = "test@system.co.il"

    #Admin Alert Email
    AdminAlertEmail = "test@system.co.il"

 }
}

function Save-LocalConfig {
    try {
        # צור תיקייה אם לא קיימת
        if (-not (Test-Path $script:LocalConfigCacheDir)) {
            New-Item -ItemType Directory -Path $script:LocalConfigCacheDir -Force | Out-Null
        }
        
        $config = @{
            vCenterServers  = $script:vCenterServers
            DefaultSettings = $script:DefaultSettings
            LastSaved       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            SavedBy         = "$env:USERDOMAIN\$env:USERNAME"
        }
        
        $config | ConvertTo-Json -Depth 10 | Out-File -FilePath $script:LocalConfigCachePath -Encoding UTF8 -Force
        
        #Write-Host "Configuration saved to LOCAL file" -ForegroundColor Green
        return $true
    }
    catch {
        #Write-Host "Failed to save local config: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Show-ResetConfigDialog {
    #Write-Host "Network path: $script:SharedConfigPath" -ForegroundColor Cyan
    #Write-Host "Network exists: $(Test-Path $script:SharedConfigPath)" -ForegroundColor Cyan
    #Write-Host "Local path: $script:LocalConfigCachePath" -ForegroundColor Cyan
    #Write-Host "Local exists: $(Test-Path $script:LocalConfigCachePath)" -ForegroundColor Cyan

    $resetForm = New-Object System.Windows.Forms.Form
    $resetForm.Text = "Reset Configuration"
    $resetForm.Size = New-Object System.Drawing.Size(450, 220)
    $resetForm.StartPosition = "CenterParent"
    $resetForm.FormBorderStyle = "FixedDialog"
    $resetForm.MaximizeBox = $false
    $resetForm.MinimizeBox = $false
    $resetForm.BackColor = [System.Drawing.Color]::FromArgb(54, 54, 68)
    $resetForm.TopMost = $true
    
    $lblQuestion = New-Object System.Windows.Forms.Label
    $lblQuestion.Location = New-Object System.Drawing.Point(20, 20)
$lblQuestion.Size = New-Object System.Drawing.Size(380, 30)
    $lblQuestion.Text = "Select configuration source to reset from:"
    $lblQuestion.Font = New-Object System.Drawing.Font("Segoe UI", 11)
    $lblQuestion.ForeColor = [System.Drawing.Color]::FromArgb(238, 240, 244)
    $resetForm.Controls.Add($lblQuestion)
    
    $btnNetwork = New-Object System.Windows.Forms.Button
    $btnNetwork.Location = New-Object System.Drawing.Point(20, 60)
    $btnNetwork.Size = New-Object System.Drawing.Size(180, 50)
    $btnNetwork.Text = "Network (Original)"
    $btnNetwork.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnNetwork.FlatStyle = "Flat"
    $btnNetwork.BackColor = [System.Drawing.Color]::FromArgb(45, 135, 183)
    $btnNetwork.ForeColor = [System.Drawing.Color]::White
    $btnNetwork.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnNetwork.Enabled = [bool](Wait-JsonFileAvailable -Path $script:SharedConfigPath -TimeoutMs $script:SharedConfigTimeoutMs)
    $btnNetwork.Add_Click({
        $resetForm.Tag = "Network"
        $resetForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $resetForm.Close()
    })
    $resetForm.Controls.Add($btnNetwork)
    
    $btnLocal = New-Object System.Windows.Forms.Button
    $btnLocal.Location = New-Object System.Drawing.Point(210, 60)
    $btnLocal.Size = New-Object System.Drawing.Size(180, 50)
    $btnLocal.Text = "Local (My Settings)"
    $btnLocal.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $btnLocal.FlatStyle = "Flat"
    $btnLocal.BackColor = [System.Drawing.Color]::FromArgb(62, 178, 122)
    $btnLocal.ForeColor = [System.Drawing.Color]::White
    $btnLocal.Cursor = [System.Windows.Forms.Cursors]::Hand
    $btnLocal.Enabled = [bool](Wait-JsonFileAvailable -Path $script:LocalConfigCachePath)
    $btnLocal.Add_Click({
        $resetForm.Tag = "Local"
        $resetForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $resetForm.Close()
    })
    $resetForm.Controls.Add($btnLocal)
    
    $lblNote = New-Object System.Windows.Forms.Label
    $lblNote.Location = New-Object System.Drawing.Point(20, 120)
$lblNote.Size = New-Object System.Drawing.Size(420, 20)
    $lblNote.Text = "Network = Original settings | Local = Your saved settings"
    $lblNote.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $lblNote.ForeColor = [System.Drawing.Color]::FromArgb(188, 192, 202)
    $resetForm.Controls.Add($lblNote)
    
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(160, 145)
    $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
    $btnCancel.Text = "Cancel"
    $btnCancel.FlatStyle = "Flat"
    $btnCancel.Add_Click({
        $resetForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $resetForm.Close()
    })
    $resetForm.Controls.Add($btnCancel)
    
    $result = $resetForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $resetForm.Tag
    }
    return $null
}

function Reset-Configuration {
    $source = Show-ResetConfigDialog
    
    if (-not $source) { return $false }
    
    $configLoaded = $false
    
    if ($source -eq "Network") {
        $networkConfig = Load-ConfigFromFile -Path $script:SharedConfigPath -TimeoutMs $script:SharedConfigTimeoutMs
        if ($networkConfig) {
            $configLoaded = Apply-Config -Config $networkConfig
            if ($configLoaded) {
                Save-LocalConfig | Out-Null
            }
        }
    }
    else {
        $localConfig = Load-ConfigFromFile -Path $script:LocalConfigCachePath
        if ($localConfig) {
            $configLoaded = Apply-Config -Config $localConfig
        }
    }
    
    if ($configLoaded) {
        [System.Windows.Forms.MessageBox]::Show(
            "Configuration reset from $source successfully!",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return $true
    }
    else {
        Write-Log "Failed to reset configuration from $source." -Level "ERROR"
        return $false
    }
}

function Refresh-SettingsForm {
    try {
        # SCCM Settings
        if ($txtSettingsSCCM) { $txtSettingsSCCM.Text = $script:DefaultSettings.SCCMServer }
        if ($txtSettingsSiteCode) { $txtSettingsSiteCode.Text = $script:DefaultSettings.SCCMSiteCode }
        
        # DNS Settings
        if ($txtSettingsDNS) { $txtSettingsDNS.Text = $script:DefaultSettings.DNSServer }
        if ($txtSettingsDNSZone) { $txtSettingsDNSZone.Text = $script:DefaultSettings.DNSZone }
        
        # AD Settings
        if ($txtSettingsDC) { $txtSettingsDC.Text = $script:DefaultSettings.DomainController }
        if ($txtSettingsEmail) { $txtSettingsEmail.Text = $script:DefaultSettings.EmailNotification }
        
        # General Options
        if ($chkEnableLogging) { $chkEnableLogging.Checked = $script:DefaultSettings.EnableLogging }
        if ($chkConfirmDeletion) { $chkConfirmDeletion.Checked = $script:DefaultSettings.ConfirmDeletion }
        if ($chkDeleteFromDisk) { $chkDeleteFromDisk.Checked = $script:DefaultSettings.DeleteFromDisk }
        if ($numTimeout) { $numTimeout.Value = $script:DefaultSettings.PowerOffTimeout }
        if ($txtLogPath) { $txtLogPath.Text = $script:DefaultSettings.LogPath }
        
        # Email Settings
        if ($chkEnableEmail) { $chkEnableEmail.Checked = $script:DefaultSettings.EnableEmailNotification }
        if ($txtSMTPServer) { $txtSMTPServer.Text = $script:DefaultSettings.SMTPServer }
        if ($numSMTPPort) { $numSMTPPort.Value = $script:DefaultSettings.SMTPPort }
        if ($chkSMTPSSL) { $chkSMTPSSL.Checked = $script:DefaultSettings.SMTPUseSSL }
        if ($txtSMTPFrom) { $txtSMTPFrom.Text = $script:DefaultSettings.SMTPFrom }
        
        # Email Groups - Common
        if ($script:lstGroupsCommon) {
            $script:lstGroupsCommon.Items.Clear()
            if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsCommon)) {
                $groups = $script:DefaultSettings.GroupsCommon.Split(',')
                foreach ($group in $groups) {
                    $g = $group.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($g)) {
                        [void]$script:lstGroupsCommon.Items.Add($g)
                    }
                }
            }
        }
        
        # Email Groups - Windows
        if ($script:lstGroupsWindows) {
            $script:lstGroupsWindows.Items.Clear()
            if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsWindows)) {
                $groups = $script:DefaultSettings.GroupsWindows.Split(',')
                foreach ($group in $groups) {
                    $g = $group.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($g)) {
                        [void]$script:lstGroupsWindows.Items.Add($g)
                    }
                }
            }
        }
        
        # Email Groups - Linux
        if ($script:lstGroupsLinux) {
            $script:lstGroupsLinux.Items.Clear()
            if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsLinux)) {
                $groups = $script:DefaultSettings.GroupsLinux.Split(',')
                foreach ($group in $groups) {
                    $g = $group.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($g)) {
                        [void]$script:lstGroupsLinux.Items.Add($g)
                    }
                }
            }
        }
        
        # Status
        if ($lblSettingsStatus) {
            $lblSettingsStatus.Text = "Settings refreshed from $($script:ConfigSource)"
            $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        }
        
        #Write-Host "Settings form refreshed successfully" -ForegroundColor Green
    }
    catch {
        #Write-Host "Error refreshing form: $($_.Exception.Message)" -ForegroundColor Red
        if ($lblSettingsStatus) {
            $lblSettingsStatus.Text = "Error refreshing settings"
            $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
        }
    }
}

<#Check Permissions
if (-not (CheckPermission)){
    $wshell = New-Object -ComObject Wscript.Shell
    $result = $wshell.Popup("You're Not Authorized!", 0,"Warnning", 0)
    exit
}
#>

# ============================================
# Initialize Configuration on Startup
# ============================================
Initialize-Configuration | Out-Null

#endregion

#region Email Functions
function Send-DeletionSummaryReport {
    param(
        [array]$SuccessList,
        [array]$FailedList,
        [array]$SkippedList,
        [string]$Reason
    )
    
    try {
        $totalVMs = $SuccessList.Count + $FailedList.Count + $SkippedList.Count
        
        # אם אין כלום - אל תשלח מייל
        if ($totalVMs -eq 0) {
            Write-Log "No VMs to report - skipping email"
            return
        }
        
        # בניית HTML
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            margin: 20px; 
            background-color: #f5f5f5; 
        }
        .container { 
            max-width: 1000px; 
            margin: 0 auto; 
            background: white; 
            padding: 30px; 
            border-radius: 10px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.1); 
        }
        h1 { 
            color: #333; 
            border-bottom: 3px solid #0066cc; 
            padding-bottom: 15px; 
            margin-bottom: 20px;
        }
        .summary { 
            display: flex; 
            gap: 20px; 
            margin: 25px 0; 
            flex-wrap: wrap;
        }
        .summary-box { 
            padding: 20px 25px; 
            border-radius: 10px; 
            color: white; 
            text-align: center; 
            flex: 1; 
            min-width: 150px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        }
        .summary-box.success { background: linear-gradient(135deg, #28a745 0%, #20c997 100%); }
        .summary-box.failed { background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); }
        .summary-box.skipped { background: linear-gradient(135deg, #6c757d 0%, #5a6268 100%); }
        .summary-box h2 { margin: 0; font-size: 36px; font-weight: bold; }
        .summary-box p { margin: 8px 0 0 0; font-size: 14px; opacity: 0.95; }
        
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-top: 20px; 
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }
        th, td { 
            border: 1px solid #dee2e6; 
            padding: 12px; 
            text-align: left; 
        }
        th { 
            background-color: #f8f9fa; 
            font-weight: 600; 
            color: #495057;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        tr:hover { background-color: #f8f9fa; }
        
        .status-icon { 
            font-size: 16px; 
            font-weight: bold; 
            text-align: center;
        }
        .status-success { color: #28a745; }
        .status-failed { color: #dc3545; }
        .status-na { color: #adb5bd; }
        
        .errors { 
            color: #dc3545; 
            font-size: 12px; 
            line-height: 1.5;
        }
        .info { 
            background: linear-gradient(135deg, #e7f3ff 0%, #cfe2ff 100%); 
            padding: 15px 20px; 
            border-radius: 8px; 
            margin: 20px 0; 
            border-left: 4px solid #0066cc;
        }
        .info strong { color: #0066cc; }
        
        h2 { 
            margin-top: 30px; 
            padding: 10px 0; 
            border-bottom: 2px solid #e9ecef;
        }
        
        .vm-name { 
            font-weight: 600; 
            color: #212529;
        }
        
        .footer {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 2px solid #e9ecef;
            text-align: center;
            color: #6c757d;
            font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🗑️ VM Deletion Summary Report</h1>
        
        <div class="info">
            <strong>👤 Initiated by:</strong> $env:USERDOMAIN\$env:USERNAME<br>
            <strong>📅 Date:</strong> $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')<br>
            <strong>🖥️ vCenter:</strong> $($global:DefaultVIServer.Name)<br>
            <strong>📝 Reason:</strong> $Reason
        </div>
        
        <div class="summary">
"@

        # הצג רק קופסאות עם ערכים
        if ($SuccessList.Count -gt 0) {
            $html += @"
            <div class="summary-box success">
                <h2>$($SuccessList.Count)</h2>
                <p>✅ Successful</p>
            </div>
"@
        }
        
        if ($FailedList.Count -gt 0) {
            $html += @"
            <div class="summary-box failed">
                <h2>$($FailedList.Count)</h2>
                <p>❌ Failed</p>
            </div>
"@
        }
        
        if ($SkippedList.Count -gt 0) {
            $html += @"
            <div class="summary-box skipped">
                <h2>$($SkippedList.Count)</h2>
                <p>⏭️ Skipped</p>
            </div>
"@
        }

        $html += @"
        </div>
"@

        # טבלת כשלונות - רק אם יש
        if ($FailedList.Count -gt 0) {
            $html += @"
        <h2 style="color: #dc3545;">❌ Failed VMs ($($FailedList.Count))</h2>
        <table>
            <tr>
                <th>VM Name</th>
                <th>OS Type</th>
                <th style="text-align: center;">SCCM</th>
                <th style="text-align: center;">Active Directory</th>
                <th style="text-align: center;">DNS</th>
                <th style="text-align: center;">VMware</th>
                <th>Error Details</th>
            </tr>
"@
            foreach ($vm in $FailedList) {
                # בדוק שיש שם VM
                if ([string]::IsNullOrWhiteSpace($vm.VMName)) {
                    continue
                }
                
                $sccmStatus = if ($vm.SCCMDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } else { 
                    '<span class="status-icon status-failed">✘</span>' 
                }
                
                $adStatus = if ($vm.ADDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } else { 
                    '<span class="status-icon status-failed">✘</span>' 
                }
                
                $dnsStatus = if ($vm.DNSDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } else { 
                    '<span class="status-icon status-failed">✘</span>' 
                }
                
                $vmStatus = if ($vm.VMDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } else { 
                    '<span class="status-icon status-failed">✘</span>' 
                }
                
                $errors = if ($vm.Errors -and $vm.Errors.Count -gt 0) { 
                    "• " + ($vm.Errors -join "<br>• ") 
                } else { 
                    "-" 
                }
                
                $osType = if ([string]::IsNullOrWhiteSpace($vm.OSType)) { "Unknown" } else { $vm.OSType }
                
                $html += @"
            <tr>
                <td class="vm-name">$($vm.VMName)</td>
                <td>$osType</td>
                <td style="text-align: center;">$sccmStatus</td>
                <td style="text-align: center;">$adStatus</td>
                <td style="text-align: center;">$dnsStatus</td>
                <td style="text-align: center;">$vmStatus</td>
                <td class="errors">$errors</td>
            </tr>
"@
            }
            $html += "</table>"
        }

        # טבלת הצלחות - רק אם יש
        if ($SuccessList.Count -gt 0) {
            $html += @"
        <h2 style="color: #28a745;">✅ Successful VMs ($($SuccessList.Count))</h2>
        <table>
            <tr>
                <th>VM Name</th>
                <th>OS Type</th>
                <th style="text-align: center;">SCCM</th>
                <th style="text-align: center;">Active Directory</th>
                <th style="text-align: center;">DNS</th>
                <th style="text-align: center;">VMware</th>
                <th style="text-align: center;">Archived</th>
            </tr>
"@
            foreach ($vm in $SuccessList) {
                # בדוק שיש שם VM
                if ([string]::IsNullOrWhiteSpace($vm.VMName)) {
                    continue
                }
                
                $sccmStatus = if ($vm.SCCMDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } elseif ($vm.OSType -eq "Linux") {
                    '<span class="status-icon status-na">N/A</span>'
                } else { 
                    '<span class="status-icon status-na">-</span>' 
                }
                
                $adStatus = if ($vm.ADDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } elseif ($vm.OSType -eq "Linux") {
                    '<span class="status-icon status-na">N/A</span>'
                } else { 
                    '<span class="status-icon status-na">-</span>' 
                }
                
                $dnsStatus = if ($vm.DNSDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } else { 
                    '<span class="status-icon status-na">-</span>' 
                }
                
                $vmStatus = if ($vm.VMDeleted) { 
                    '<span class="status-icon status-success">✔</span>' 
                } else { 
                    '<span class="status-icon status-na">-</span>' 
                }
                
                $archiveStatus = if ($vm.Archived) { 
                    '<span class="status-icon status-success">✔</span>' 
                } else { 
                    '<span class="status-icon status-na">-</span>' 
                }
                
                $osType = if ([string]::IsNullOrWhiteSpace($vm.OSType)) { "Unknown" } else { $vm.OSType }
                
                $html += @"
            <tr>
                <td class="vm-name">$($vm.VMName)</td>
                <td>$osType</td>
                <td style="text-align: center;">$sccmStatus</td>
                <td style="text-align: center;">$adStatus</td>
                <td style="text-align: center;">$dnsStatus</td>
                <td style="text-align: center;">$vmStatus</td>
                <td style="text-align: center;">$archiveStatus</td>
            </tr>
"@
            }
            $html += "</table>"
        }

        # טבלת מדולגים - רק אם יש
        if ($SkippedList.Count -gt 0) {
            $html += @"
        <h2 style="color: #6c757d;">⏭️ Skipped VMs ($($SkippedList.Count))</h2>
        <table>
            <tr>
                <th>VM Name</th>
                <th>OS Type</th>
                <th>Skip Reason</th>
            </tr>
"@
            foreach ($vm in $SkippedList) {
                # בדוק שיש שם VM
                if ([string]::IsNullOrWhiteSpace($vm.VMName)) {
                    continue
                }
                
                $skipReason = if ([string]::IsNullOrWhiteSpace($vm.SkipReason)) { "Unknown" } else { $vm.SkipReason }
                $osType = if ([string]::IsNullOrWhiteSpace($vm.OSType)) { "Unknown" } else { $vm.OSType }
                
                $html += @"
            <tr>
                <td class="vm-name">$($vm.VMName)</td>
                <td>$osType</td>
                <td>$skipReason</td>
            </tr>
"@
            }
            $html += "</table>"
        }

        $html += @"
        <div class="footer">
            Generated by VM Deletion Tool | $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
        </div>
    </div>
</body>
</html>
"@

        # שלח מייל
        $subjectParts = @()
        if ($SuccessList.Count -gt 0) { $subjectParts += "$($SuccessList.Count) Success" }
        if ($FailedList.Count -gt 0) { $subjectParts += "$($FailedList.Count) Failed" }
        if ($SkippedList.Count -gt 0) { $subjectParts += "$($SkippedList.Count) Skipped" }
        
        $subject = "$($script:DefaultSettings.EmailSubjectPrefix) Summary: " + ($subjectParts -join ", ")
        
        $emailParams = @{
            From       = $script:DefaultSettings.SMTPFrom
            To         = $script:DefaultSettings.AdminAlertEmail
            Subject    = $subject
            Body       = $html
            BodyAsHtml = $true
            SmtpServer = $script:DefaultSettings.SMTPServer
            Port       = $script:DefaultSettings.SMTPPort
            Encoding   = [System.Text.Encoding]::UTF8
        }
        
        if ($script:DefaultSettings.SMTPUseSSL) {
            $emailParams.UseSsl = $true
        }
        
        Send-MailMessage @emailParams -ErrorAction Stop
        
        Write-Log "Deletion summary report sent to $($script:DefaultSettings.AdminAlertEmail)"
    }
    catch {
        Write-Log "Failed to send summary report: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Get-ADUserEmail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SamAccountName,
        [string]$DomainController
    )
    
    $result = @{
        Success = $false
        Email   = ""
        Message = ""
    }
    
    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            $result.Message = "ActiveDirectory module not installed"
            return $result
        }
        
        Import-Module ActiveDirectory -ErrorAction Stop
        
        $adParams = @{
            Identity    = $SamAccountName
            Properties  = @("mail", "EmailAddress")
            ErrorAction = 'Stop'
        }
        
        if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
            $adParams.Server = $DomainController
        }
        
        $user = Get-ADUser @adParams
        
        if ($user) {
            $email = if ($user.mail) { $user.mail } else { $user.EmailAddress }
            if ($email) {
                $result.Success = $true
                $result.Email = $email
                $result.Message = "Email found"
            }
            else {
                $result.Message = "User found but no email address"
            }
        }
    }
    catch {
        $result.Message = "Error: $($_.Exception.Message)"
    }
    
    return $result
}

function Get-ADGroupMemberEmails {
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupName,
        [string]$DomainController
    )
    
    $emails = @()
    
    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            Write-Log "ActiveDirectory module not installed" -Level "WARNING"
            return $emails
        }
        
        Import-Module ActiveDirectory -ErrorAction Stop
        
        $adParams = @{
            Identity    = $GroupName
            Recursive   = $true
            ErrorAction = 'Stop'
        }
        
        if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
            $adParams.Server = $DomainController
        }
        
        $members = Get-ADGroupMember @adParams
        
        foreach ($member in $members) {
            if ($member.objectClass -eq "user") {
                $userParams = @{
                    Identity    = $member.SamAccountName
                    Properties  = @("mail", "EmailAddress")
                    ErrorAction = 'SilentlyContinue'
                }
                
                if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
                    $userParams.Server = $DomainController
                }
                
                $user = Get-ADUser @userParams
                
                if ($user) {
                    $email = if ($user.mail) { $user.mail } else { $user.EmailAddress }
                    if ($email -and $emails -notcontains $email) {
                        $emails += $email
                    }
                }
            }
        }
    }
    catch {
        Write-Log "Error getting group members for $GroupName : $($_.Exception.Message)" -Level "WARNING"
    }
    
    return $emails
}

function Send-VMDeletionNotification {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$VMDetails,
        [string]$Reason,
        [string]$OSType,
        [bool]$DeletionSuccess
    )
    
    $result = @{
        Success = $false
        Message = ""
        SentTo  = @()
    }
    
    if (-not $script:DefaultSettings.EnableEmailNotification) {
        $result.Message = "Email notifications disabled"
        return $result
    }
    
    if ([string]::IsNullOrWhiteSpace($script:DefaultSettings.SMTPServer)) {
        $result.Message = "SMTP server not configured"
        return $result
    }
    
    try {
        $dc = $script:DefaultSettings.DomainController
        $allRecipients = @()
        
        # ========== פונקציה פנימית לעיבוד נמענים (קבוצות או משתמשים) ==========
        function Get-RecipientsFromList {
            param(
                [string]$RecipientList,
                [string]$DomainController,
                [string]$ListType
            )
            
            $emails = @()
            if ([string]::IsNullOrWhiteSpace($RecipientList)) { return $emails }
            
            $items = $RecipientList.Split(',')
            foreach ($item in $items) {
                $item = $item.Trim()
                if ([string]::IsNullOrWhiteSpace($item)) { continue }
                
                # בדוק אם זו כתובת אימייל ישירה
                if ($item -match '@') {
                    Write-Log "Adding direct email ($ListType): $item"
                    $emails += $item
                }
                # בדוק אם זה משתמש AD (לפי פורמט domain\user או user)
                elseif ($item -match '\\' -or $item -notmatch '\s') {
                    # נסה קודם כקבוצה
                    try {
                        $groupEmails = Get-ADGroupMemberEmails -GroupName $item -DomainController $DomainController -ErrorAction Stop
                        if ($groupEmails.Count -gt 0) {
                            Write-Log "Found $($groupEmails.Count) emails in group ($ListType): $item"
                            $emails += $groupEmails
                            continue
                        }
                    }
                    catch {
                        # לא קבוצה, נסה כמשתמש
                    }
                    
                    # נסה כמשתמש בודד
                    $userEmail = Get-ADUserEmail -SamAccountName $item -DomainController $DomainController
                    if ($userEmail.Success -and -not [string]::IsNullOrWhiteSpace($userEmail.Email)) {
                        Write-Log "Found user email ($ListType): $item -> $($userEmail.Email)"
                        $emails += $userEmail.Email
                    }
                    else {
                        Write-Log "Could not find email for: $item - $($userEmail.Message)" -Level "WARNING"
                    }
                }
            }
            
            return $emails
        }
        
        # ========== בדיקה אם זה Template ==========
        $vmData = $script:allVMsData | Where-Object { $_.Name -eq $VMDetails.VMName } | Select-Object -First 1
        $isTemplate = if ($vmData) { $vmData.IsTemplate } else { $false }
        
        # ========== בדיקה אם Linux מ-Clusters ספציפיים ==========
        # cluster1, cluster2, cluster3 - שולחים לכולם: Linux + Common + Windows
        $linuxSpecialClusters = @("LINUXGPU", "LINUXHANA", "LINUXNOORA" ,"LINUXORASYN")
        $isLinuxSpecialCluster = $false
        if ($OSType -eq "Linux" -and $VMDetails.ClusterName) {
            $clusterName = $VMDetails.ClusterName.ToLower()
            foreach ($specialCluster in $linuxSpecialClusters) {
                if ($clusterName -like "*$specialCluster*") {
                    $isLinuxSpecialCluster = $true
                    Write-Log "Linux VM from special cluster: $($VMDetails.ClusterName) - sending to all teams"
                    break
                }
            }
        }
        
        if ($isTemplate) {
            # ========== Templates - רק Windows ==========
            Write-Log "Template detected - sending only to Windows group"
            $windowsEmails = Get-RecipientsFromList -RecipientList $script:DefaultSettings.GroupsWindows -DomainController $dc -ListType "Windows"
            $allRecipients += $windowsEmails
        }
        elseif ($isLinuxSpecialCluster) {
            # ========== Linux מ-Clusters מיוחדים - שולחים לכולם ==========
            Write-Log "Linux from special cluster - sending to Linux + Common + Windows"
            
            $commonEmails = Get-RecipientsFromList -RecipientList $script:DefaultSettings.GroupsCommon -DomainController $dc -ListType "Common"
            $allRecipients += $commonEmails
            
            $windowsEmails = Get-RecipientsFromList -RecipientList $script:DefaultSettings.GroupsWindows -DomainController $dc -ListType "Windows"
            $allRecipients += $windowsEmails
            
            $linuxEmails = Get-RecipientsFromList -RecipientList $script:DefaultSettings.GroupsLinux -DomainController $dc -ListType "Linux"
            $allRecipients += $linuxEmails
        }
        elseif ($OSType -eq "Linux") {
            # ========== Linux רגיל (לא מ-Clusters מיוחדים) - לא שולחים מייל ==========
            Write-Log "Linux VM not from special cluster - no email will be sent to Linux team"
            $commonEmails = Get-RecipientsFromList -RecipientList $script:DefaultSettings.GroupsCommon -DomainController $dc -ListType "Common"
            $allRecipients += $commonEmails
        }
        else {
            # ========== Windows VMs - Common + Windows ==========
            $commonEmails = Get-RecipientsFromList -RecipientList $script:DefaultSettings.GroupsCommon -DomainController $dc -ListType "Common"
            $allRecipients += $commonEmails
            
            $windowsEmails = Get-RecipientsFromList -RecipientList $script:DefaultSettings.GroupsWindows -DomainController $dc -ListType "Windows"
            $allRecipients += $windowsEmails
        }
        
        # ========== אחראי השרת מ-Annotation (srvOWN1) ==========
        if (-not [string]::IsNullOrWhiteSpace($VMDetails.ServerOwner)) {
            Write-Log "Getting email for server owner: $($VMDetails.ServerOwner)"
            $ownerEmail = Get-ADUserEmail -SamAccountName $VMDetails.ServerOwner -DomainController $dc
            if ($ownerEmail.Success -and -not [string]::IsNullOrWhiteSpace($ownerEmail.Email)) {
                $allRecipients += $ownerEmail.Email
                Write-Log "Found owner email: $($ownerEmail.Email)"
            }
            else {
                Write-Log "Could not find email for owner: $($VMDetails.ServerOwner) - $($ownerEmail.Message)" -Level "WARNING"
            }
        }
        
        # ========== הסר כפילויות וערכים ריקים ==========
        $allRecipients = $allRecipients | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
        
        if ($allRecipients.Count -eq 0) {
            $result.Message = "No recipients found"
            Write-Log "No email recipients found for $($VMDetails.VMName)" -Level "WARNING"
            return $result
        }
        
        Write-Log "Sending email to $($allRecipients.Count) recipient(s): $($allRecipients -join ', ')"

        # ========== בניית גוף המייל ==========
        $status = if ($DeletionSuccess) { "Successfully Deleted" } else { "Deletion Failed" }
        $statusColor = if ($DeletionSuccess) { "#28a745" } else { "#dc3545" }
        $headerColor = if ($DeletionSuccess) { "#0066cc" } else { "#dc3545" }
        
        $body = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            margin: 0; 
            padding: 0; 
            background-color: #f5f5f5;
        }
        .container {
            max-width: 700px;
            margin: 20px auto;
            background-color: #ffffff;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        .header { 
            background-color: $headerColor; 
            color: white; 
            padding: 20px 25px;
        }
        .header h2 {
            margin: 0;
            font-size: 24px;
        }
        .header p {
            margin: 5px 0 0 0;
            opacity: 0.9;
            font-size: 14px;
        }
        .content { 
            padding: 25px;
        }
        .status-badge { 
            display: inline-block;
            font-size: 16px; 
            font-weight: bold; 
            color: white;
            background-color: $statusColor;
            padding: 8px 16px;
            border-radius: 4px;
            margin-bottom: 20px;
        }
        .vm-name {
            font-size: 22px;
            font-weight: bold;
            color: #333;
            margin-bottom: 20px;
        }
        table { 
            border-collapse: collapse; 
            width: 100%; 
            margin-top: 15px;
            font-size: 14px;
        }
        th, td { 
            border: 1px solid #e0e0e0; 
            padding: 12px 15px; 
            text-align: left; 
        }
        th { 
            background-color: #f8f9fa; 
            font-weight: 600;
            color: #495057;
            width: 35%;
        }
        tr:nth-child(even) {
            background-color: #fafafa;
        }
        .os-windows {
            color: #0078d4;
            font-weight: bold;
        }
        .os-linux {
            color: #cc6600;
            font-weight: bold;
        }
        .reason-cell {
            background-color: #fff3cd;
            font-weight: 500;
        }
        .footer { 
            margin-top: 25px; 
            padding: 20px 25px; 
            background-color: #f8f9fa;
            border-top: 1px solid #e0e0e0; 
            font-size: 12px; 
            color: #6c757d;
        }
        .footer p {
            margin: 5px 0;
        }
        .warning-box {
            background-color: #fff3cd;
            border: 1px solid #ffc107;
            border-radius: 4px;
            padding: 12px 15px;
            margin-top: 20px;
            font-size: 13px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>🖥️ VM Deletion Notification</h2>
            <p>Automated notification from VM Deletion Tool</p>
        </div>
        <div class="content">
            <span class="status-badge">$status</span>
            
            <div class="vm-name">$($VMDetails.VMName)</div>
            
            <table>
                <tr>
                    <th>Property</th>
                    <th>Value</th>
                </tr>
                <tr>
                    <td>vCenter Server</td>
                    <td>$($VMDetails.VCenterName)</td>
                </tr>
                <tr>
                    <td>Cluster</td>
                    <td>$($VMDetails.ClusterName)</td>
                </tr>
                <tr>
                    <td>Folder</td>
                    <td>$($VMDetails.FolderName)</td>
                </tr>
                <tr>
                    <td>Operating System</td>
                    <td><span class="os-$($OSType.ToLower())">$OSType</span> - $($VMDetails.OperatingSystem)</td>
                </tr>
                <tr>
                    <td>Server Owner</td>
                    <td>$($VMDetails.ServerOwner)</td>
                </tr>
                <tr>
                    <td>CPU Cores</td>
                    <td>$($VMDetails.TotalCPUs) vCPU</td>
                </tr>
                <tr>
                    <td>Memory</td>
                    <td>$($VMDetails.TotalMemoryGB) GB</td>
                </tr>
                <tr>
                    <td>Total Disk Size</td>
                    <td>$($VMDetails.TotalDisksGB) GB</td>
                </tr>
                <tr>
                    <td>Power State (Before Deletion)</td>
                    <td>$($VMDetails.PowerState)</td>
                </tr>
                <tr>
                    <td>Deletion Reason</td>
                    <td class="reason-cell">$Reason</td>
                </tr>
                <tr>
                    <td>Deleted By</td>
                    <td><strong>$env:USERDOMAIN\$env:USERNAME</strong></td>
                </tr>
                <tr>
                    <td>Deletion Time</td>
                    <td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td>
                </tr>
            </table>
            
            <div class="warning-box">
                ⚠️ <strong>Note:</strong> This VM has been permanently deleted. 
                Please update any relevant documentation, monitoring systems, and backup schedules accordingly.
            </div>
        </div>
        <div class="footer">
            <p>📧 This is an automated notification. Please do not reply to this email.</p>
            <p>🔧 VM Deletion Tool v2.0.0</p>
            <p>📅 Sent at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>
    </div>
</body>
</html>
"@
        
        # ========== פרמטרים לשליחת מייל ==========
        $subject = "$($script:DefaultSettings.EmailSubjectPrefix) $($VMDetails.VMName) - $status"
        
        $mailParams = @{
            From       = $script:DefaultSettings.SMTPFrom
            To         = $allRecipients
            Subject    = $subject
            Body       = $body
            BodyAsHtml = $true
            SmtpServer = $script:DefaultSettings.SMTPServer
            Port       = $script:DefaultSettings.SMTPPort
            Priority   = "High"
            Encoding   = [System.Text.Encoding]::UTF8
        }
        
        if ($script:DefaultSettings.SMTPUseSSL) {
            $mailParams.UseSsl = $true
        }
        
        # ========== שליחת המייל ==========
        Send-MailMessage @mailParams -ErrorAction Stop
        
        $result.Success = $true
        $result.SentTo = $allRecipients
        $result.Message = "Email sent to $($allRecipients.Count) recipient(s)"
        
        Write-Log "Email notification sent successfully for $($VMDetails.VMName) to: $($allRecipients -join ', ')"
    }
    catch {
        $result.Message = "Email error: $($_.Exception.Message)"
        Write-Log "Email notification failed for $($VMDetails.VMName): $($_.Exception.Message)" -Level "ERROR"
    }
    
    return $result
}

function Send-FailureEmail {
    param(
        [string]$Subject,
        [string]$Body
    )

    try {
        $smtpServer = $script:DefaultSettings.SmtpServer
        $from       = $script:DefaultSettings.EmailNotification
        $to         = $script:DefaultSettings.AdminAlertEmail.Split(";")

        Send-MailMessage `
            -From $from `
            -To $to `
            -Subject $Subject `
            -Body $Body `
            -SmtpServer $smtpServer `
            -ErrorAction Stop

        Write-Log "Failure email sent successfully"
    }
    catch {
        Write-Log "FAILED sending mail: $($_.Exception.Message)" -Level "ERROR"
    }
}

#endregion Email Functions

# Main Tab Events

function Get-ShortTaskFailureMessage {
    param([int]$TaskIndex)

    $component = switch ($TaskIndex) {
        0 { "vCenter" }
        1 { "SCCM" }
        2 { "AD" }
        3 { "DNS" }
        4 { "VM" }
        5 { "Archive" }
        6 { "Email" }
        default { "Task" }
    }

    return "Failed - $component failed. See LOGS."
}
function Update-TaskStatus {
    param([int]$TaskIndex, [string]$Status, [string]$Message)
    $sl = $statusLabels[$TaskIndex]
    $rl = $resultLabels[$TaskIndex]
    switch ($Status) {
        "Running" {
            $sl.Text = "RUN"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(20,184,166)
            $rl.Text = "In Progress..."
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(20,184,166)
        }
        "Success" {
            $sl.Text = "OK"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(34,197,94)
            $rl.Text = $Message
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(34,197,94)
        }
        "Failed" {
            $sl.Text = "X"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(225,29,72)
            $rl.Text = Get-ShortTaskFailureMessage -TaskIndex $TaskIndex
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(225,29,72)
        }
        "Skipped" {
            $sl.Text = "SKIP"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(234,179,8)
            $rl.Text = $Message
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(234,179,8)
        }
        "Pending" {
            $sl.Text = "WAIT"
            $sl.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225)
            $rl.Text = "Pending"
            $rl.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225)
        }
    }
    if ($form) { $form.Refresh() }
    [System.Windows.Forms.Application]::DoEvents()
}

function Reset-AllStatus { 
    for ($i = 0; $i -lt $tasks.Count; $i++) { 
        Update-TaskStatus -TaskIndex $i -Status "Pending" -Message "Pending" 
    }
    $progressBar.Value = 0 
}

function Get-VMList {
    $script:LastAppliedVMFilterKey = ""
    $script:IsUpdatingList = $true
    
    # שמירת בחירות קיימות
    $selectedHash = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($item in $lstVMs.CheckedItems) {
        if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace($item.Text)) {
            [void]$selectedHash.Add($item.Text)
        }
    }
    # הוסף גם את הבחירות הקודמות מה-ArrayList
    if ($null -ne $script:SelectedVMNames) {
        foreach ($vmName in $script:SelectedVMNames) {
            if (-not [string]::IsNullOrWhiteSpace($vmName)) {
                [void]$selectedHash.Add($vmName)
            }
        }
    }
    
    $lstVMs.BeginUpdate()
    $lstVMs.Items.Clear()
    $script:allVMs = @()
    $script:allVMsData = @()
    $script:allTemplatesData = @()
    
    try {
        # שליפה מהירה של VMs – כולל PowerState וסטטוס Template
        $views = Get-View -ViewType VirtualMachine -Property Name,Config.GuestFullName,Config.GuestId,Runtime.PowerState,Config.Template -ErrorAction Stop
        
        if ($null -eq $views) {
            $lblVMCount.Text = "0 VMs"
            $lstVMs.EndUpdate()
            $script:IsUpdatingList = $false
            return
        }
        
        foreach ($vm in $views | Sort-Object Name) {
            # בדיקת null בסיסית
            if ($null -eq $vm -or [string]::IsNullOrWhiteSpace($vm.Name)) { 
                continue 
            }
            
            # בדוק אם זה Template (עם בדיקת null)
            $isTemplate = $false
            if ($null -ne $vm.Config) {
                $isTemplate = [bool]$vm.Config.Template
            }
            
            # דלג על Templates אם לא נבחרה האפשרות להציג אותם
            if ($isTemplate -and -not $script:IncludeTemplates) {
                continue
            }
            
            # קבל ערכים עם בדיקות null
            $osFullName = $null
            $guestId = $null
            if ($null -ne $vm.Config) {
                $osFullName = $vm.Config.GuestFullName
                $guestId = $vm.Config.GuestId
            }
            
            # קבע PowerState עם בדיקת null
            $powerState = "Unknown"
            if ($isTemplate) {
                $powerState = "Template"
            }
            elseif ($null -ne $vm.Runtime -and $null -ne $vm.Runtime.PowerState) {
                $powerState = $vm.Runtime.PowerState.ToString()
            }
            
            # קבע סוג OS
            $osType = "Other"
            $osFullNameStr = if ($osFullName) { $osFullName.ToString() } else { "" }
            $guestIdStr = if ($guestId) { $guestId.ToString() } else { "" }
            
            if ($osFullNameStr -match "Windows" -or $guestIdStr -match "windows") {
                $osType = "Windows"
            }
            elseif ($osFullNameStr -match "Linux|Ubuntu|CentOS|Red Hat|RHEL|Debian|SUSE|Fedora" -or
                    $guestIdStr -match "linux|ubuntu|centos|rhel|debian|suse") {
                $osType = "Linux"
            }
            
            # קבע OSFull
            $osFull = "Unknown"
            if (-not [string]::IsNullOrWhiteSpace($osFullNameStr)) {
                $osFull = $osFullNameStr
            }
            elseif (-not [string]::IsNullOrWhiteSpace($guestIdStr)) {
                $osFull = $guestIdStr
            }
            
            $vmData = @{
                Name       = $vm.Name
                OSType     = $osType
                OSFull     = $osFull
                PowerState = $powerState
                IsTemplate = $isTemplate
            }
            
            $script:allVMs += $vm.Name
            $script:allVMsData += $vmData
            
            if ($isTemplate) {
                $script:allTemplatesData += $vmData
            }
            
            $item = New-Object System.Windows.Forms.ListViewItem($vm.Name)
            [void]$item.SubItems.Add($powerState)
            [void]$item.SubItems.Add($osType)
            [void]$item.SubItems.Add($osFull)
            
            # צבע טקסט לפי OS או Template
            if ($isTemplate) {
                $item.ForeColor = [System.Drawing.Color]::FromArgb(196, 181, 253)  # Soft purple for templates
                $item.BackColor = [System.Drawing.Color]::FromArgb(96, 86, 132)  # Dark purple row
            }
            else {
                switch ($osType) {
                    "Windows" { $item.ForeColor = [System.Drawing.Color]::FromArgb(147, 197, 253) }
                    "Linux"   { $item.ForeColor = [System.Drawing.Color]::FromArgb(251, 191, 36) }
                    default   { $item.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225) }
                }
                
                # צבע רקע לפי PowerState
                switch ($powerState) {
                    "poweredOff" { $item.BackColor = [System.Drawing.Color]::FromArgb(104, 66, 80) }
                    "suspended"  { $item.BackColor = [System.Drawing.Color]::FromArgb(110, 94, 60) }
                    default      { $item.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 78) }
                }
            }
            
            # שחזר בחירה
            if ($selectedHash.Contains($vm.Name)) {
                $item.Checked = $true
            }
            
            [void]$lstVMs.Items.Add($item)
        }
        
        # עדכן את ה-ArrayList
        $script:SelectedVMNames.Clear()
        foreach ($name in $selectedHash) {
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                [void]$script:SelectedVMNames.Add($name)
            }
        }
        
        $templateCount = @($script:allVMsData | Where-Object { $_.IsTemplate }).Count
        $vmCount = $lstVMs.Items.Count - $templateCount
        if ($script:IncludeTemplates -and $templateCount -gt 0) {
            $lblVMCount.Text = "$vmCount VMs + $templateCount Templates"
        } else {
            $lblVMCount.Text = "$($lstVMs.Items.Count) VMs"
        }
        $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"
    }
    catch {
        Write-Log "Failed to retrieve VM list: $($_.Exception.Message); Stack: $($_.ScriptStackTrace)" -Level "ERROR"
    }
    
    $lstVMs.EndUpdate()
    $script:IsUpdatingList = $false
}

function New-VMListViewItem {
    param(
        [hashtable]$VMData,
        [System.Collections.Generic.HashSet[string]]$SelectedNames
    )

    if ($null -eq $VMData -or [string]::IsNullOrWhiteSpace($VMData.Name)) { return $null }

    $powerState = if ($VMData.PowerState) { $VMData.PowerState } else { "Unknown" }
    $osType = if ($VMData.OSType) { $VMData.OSType } else { "Other" }
    $osFull = if ($VMData.OSFull) { $VMData.OSFull } else { "Unknown" }

    $item = New-Object System.Windows.Forms.ListViewItem($VMData.Name)
    [void]$item.SubItems.Add($powerState)
    [void]$item.SubItems.Add($osType)
    [void]$item.SubItems.Add($osFull)

    if ($VMData.IsTemplate) {
        $item.ForeColor = [System.Drawing.Color]::FromArgb(196, 181, 253)
        $item.BackColor = [System.Drawing.Color]::FromArgb(96, 86, 132)
    }
    else {
        switch ($osType) {
            "Windows" { $item.ForeColor = [System.Drawing.Color]::FromArgb(147, 197, 253) }
            "Linux"   { $item.ForeColor = [System.Drawing.Color]::FromArgb(251, 191, 36) }
            default   { $item.ForeColor = [System.Drawing.Color]::FromArgb(214, 218, 225) }
        }

        switch ($powerState) {
            "poweredOff" { $item.BackColor = [System.Drawing.Color]::FromArgb(104, 66, 80) }
            "suspended"  { $item.BackColor = [System.Drawing.Color]::FromArgb(110, 94, 60) }
            default      { $item.BackColor = [System.Drawing.Color]::FromArgb(62, 62, 78) }
        }
    }

    if ($SelectedNames -and $SelectedNames.Contains($VMData.Name)) {
        $item.Checked = $true
    }

    return $item
}
function Filter-VMList {
    param(
        [string]$FilterText,
        [string]$OSFilter = "All",
        [string]$PowerFilter = "All"
    )

    if ($script:IsUpdatingList) { return }

    $filterKey = "{0}|{1}|{2}|{3}" -f $FilterText, $OSFilter, $PowerFilter, $script:IncludeTemplates
    if ($script:LastAppliedVMFilterKey -eq $filterKey) { return }

    $filterStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $script:LastAppliedVMFilterKey = $filterKey
    $script:IsUpdatingList = $true

    $selectedHash = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $lstVMs.CheckedItems) {
        if ($null -ne $item -and -not [string]::IsNullOrWhiteSpace($item.Text)) {
            [void]$selectedHash.Add($item.Text)
        }
    }
    if ($null -ne $script:SelectedVMNames) {
        foreach ($vmName in $script:SelectedVMNames) {
            if (-not [string]::IsNullOrWhiteSpace($vmName)) {
                [void]$selectedHash.Add([string]$vmName)
            }
        }
    }

    $items = New-Object 'System.Collections.Generic.List[System.Windows.Forms.ListViewItem]'
    $filterValue = if ($null -eq $FilterText) { "" } else { $FilterText.Trim() }
    $useTextFilter = -not [string]::IsNullOrWhiteSpace($filterValue)
    $useOSFilter = -not [string]::IsNullOrWhiteSpace($OSFilter) -and $OSFilter -ne "All"
    $usePowerFilter = -not [string]::IsNullOrWhiteSpace($PowerFilter) -and $PowerFilter -ne "All"

    foreach ($vmData in $script:allVMsData) {
        if ($null -eq $vmData -or [string]::IsNullOrWhiteSpace($vmData.Name)) { continue }
        if ($useTextFilter -and $vmData.Name.IndexOf($filterValue, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        if ($useOSFilter -and $vmData.OSType -ne $OSFilter) { continue }
        if ($usePowerFilter -and $vmData.PowerState -ne $PowerFilter) { continue }

        $item = New-VMListViewItem -VMData $vmData -SelectedNames $selectedHash
        if ($null -ne $item) { $items.Add($item) }
    }

    $lstVMs.BeginUpdate()
    try {
        $lstVMs.Items.Clear()
        if ($items.Count -gt 0) {
            $lstVMs.Items.AddRange($items.ToArray())
        }
    }
    finally {
        $lstVMs.EndUpdate()
        $script:IsUpdatingList = $false
    }

    $script:SelectedVMNames.Clear()
    foreach ($name in $selectedHash) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            [void]$script:SelectedVMNames.Add($name)
        }
    }

    $lblVMCount.Text = "$($lstVMs.Items.Count) of $(@($script:allVMsData).Count)"
    $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"

    $filterStopwatch.Stop()
    if ($filterStopwatch.ElapsedMilliseconds -gt 300) {
        Write-Log "VM filter completed in $($filterStopwatch.ElapsedMilliseconds) ms. Results=$($lstVMs.Items.Count), Filter='$filterValue', OS='$OSFilter', Power='$PowerFilter'"
    }
}

function Schedule-VMFilter {
    $osFilter = "All"
    if ($null -ne $cboOSFilter.SelectedItem) {
        $osFilter = $cboOSFilter.SelectedItem.ToString()
    }

    $powerFilter = "All"
    if ($null -ne $cboPowerFilter.SelectedItem) {
        $powerFilter = $cboPowerFilter.SelectedItem.ToString()
    }

    $powerFilter = switch ($powerFilter) {
        "On"        { "poweredOn" }
        "Off"       { "poweredOff" }
        "Suspended" { "suspended" }
        default     { "All" }
    }

    if ($null -eq $script:VMFilterTimer) {
        Filter-VMList -FilterText $txtFilter.Text -OSFilter $osFilter -PowerFilter $powerFilter
        return
    }

    $script:PendingVMFilterText = $txtFilter.Text
    $script:PendingVMOSFilter = $osFilter
    $script:PendingVMPowerFilter = $powerFilter
    $script:VMFilterTimer.Stop()
    $script:VMFilterTimer.Start()
}

function Set-ControlsEnabled { 
    param([bool]$Enabled)
    $lstVMs.Enabled = $true
    $lstVMs.BackColor = $uiInput
    $lstVMs.ForeColor = $uiText
    $txtFilter.Enabled = $Enabled
    $cboOSFilter.Enabled = $Enabled
    $cboPowerFilter.Enabled = $Enabled
    $btnRefresh.Enabled = $Enabled
    $btnClearSelection.Enabled = $Enabled
    $btnLoadCSV.Enabled = $Enabled
    $btnDelete.Enabled = $Enabled
    $chkSCCM.Enabled = $Enabled
    $chkAD.Enabled = $Enabled
    $chkDNS.Enabled = $Enabled
    $chkVM.Enabled = $Enabled
    $chkArchive.Enabled = $Enabled
    $chkEmail.Enabled = $Enabled
    $txtReason.Enabled = $Enabled
}

function Get-ADSearchBaseForEnvironment {
    param([string]$EnvironmentName)

    $envKey = if ([string]::IsNullOrWhiteSpace($EnvironmentName)) {
        "PROD"
    }
    else {
        $EnvironmentName.Trim().ToUpper()
    }

    try {
        $map = $script:DefaultSettings.ADSearchBaseByEnvironment
        if ($map) {
            if ($map -is [hashtable] -and $map.ContainsKey($envKey)) {
                return $map[$envKey]
            }

            if ($map.PSObject.Properties.Name -contains $envKey) {
                return $map.$envKey
            }
        }
    }
    catch { }

    return $script:DefaultSettings.ADSearchBase
}

function Resolve-VCenterEnvironment {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $script:RunEnvironment }

    foreach ($vc in $script:vCenterServers) {
        if ($vc.Name -eq $Value -or $vc.Server -eq $Value) {
            return $vc.Name
        }
    }

    return $Value
}

function Protect-Password {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingConvertToSecureStringWithPlainText', '', Justification='Plain text is accepted only transiently so it can be encrypted with ConvertFrom-SecureString.')]
    param([string]$PlainPassword)
    
    if ([string]::IsNullOrEmpty($PlainPassword)) { return "" }
    
    try {
        $secureString = ConvertTo-SecureString -String $PlainPassword -AsPlainText -Force
        $encrypted = ConvertFrom-SecureString -SecureString $secureString
        return $encrypted
    }
    catch {
        Write-ExtendedLog "Password encryption failed: $($_.Exception.Message)" -Level "ERROR" -SendEmail
        return ""
    }
}

function Unprotect-Password {
    param([string]$EncryptedPassword)
    
    if ([string]::IsNullOrEmpty($EncryptedPassword)) { return "" }
    
    try {
        $secureString = ConvertTo-SecureString -String $EncryptedPassword
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureString)
        $plainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        return $plainPassword
    }
    catch {
        Write-ExtendedLog "Password decryption failed: $($_.Exception.Message)" -Level "ERROR" -SendEmail
        return ""
    }
}

function Save-Settings {
    try {
        # Use a single directory for both configuration and logging
        $configDir = "$env:TEMP\VMDeletionTool"
        if (-not (Test-Path $configDir)) { 
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null 
        }
        
        # Update the log path to match the config directory
        $script:DefaultSettings.LogPath = $configDir
        
        $jsonConfigPath = Join-Path $configDir "config.json"
        
        # Prepare settings for saving
        $settingsToSave = @{}
        
        # Explicitly copy all settings
        foreach ($key in $script:DefaultSettings.Keys) {
            $settingsToSave[$key] = $script:DefaultSettings[$key]
        }
        
        # Encrypt sensitive data
        if (-not [string]::IsNullOrEmpty($settingsToSave.SQLPassword) -and
        (-not $settingsToSave.ContainsKey("SQLPasswordEncrypted") -or
        $settingsToSave.SQLPasswordEncrypted -ne $true)) {

        $settingsToSave.SQLPassword = Protect-Password -PlainPassword $settingsToSave.SQLPassword
        $settingsToSave["SQLPasswordEncrypted"] = $true
}
        # Serialize and save in the same structure used by startup config loading.
        $configToSave = @{
            vCenterServers  = $script:vCenterServers
            DefaultSettings = $settingsToSave
            LastSaved       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            SavedBy         = "$env:USERDOMAIN\$env:USERNAME"
        }
        $jsonSettings = $configToSave | ConvertTo-Json -Depth 10
        $jsonSettings | Out-File -FilePath $jsonConfigPath -Encoding UTF8 -Force        
        # Verify file was written
    if (Wait-JsonFileAvailable -Path $jsonConfigPath) {
        Write-ExtendedLog "Settings saved successfully to $jsonConfigPath" -Level "INFO" -Component "ConfigManager"
        return $true
        }
        else {
            Write-ExtendedLog "Failed to create settings file" -Level "ERROR" -Component "ConfigManager"
            return $false
        }
    } 
    catch { 
        Write-ExtendedLog "Save settings failed: $($_.Exception.Message)" -Level "ERROR" -Component "ConfigManager"
        Write-ExtendedLog "Exception details: $($_.Exception | Format-List * | Out-String)" -Level "ERROR" -Component "ConfigManager"
        return $false 
    }
}

function Load-Settings {
    try {
        # Use the same directory for configuration and logging
        $configDir = "$env:TEMP\VMDeletionTool"
        $jsonConfigPath = Join-Path $configDir "config.json"
        
        # Update the log path to match the config directory
        $script:DefaultSettings.LogPath = $configDir
        
        $availableJsonConfigPath = Wait-JsonFileAvailable -Path $jsonConfigPath
        if ($availableJsonConfigPath) {
            try {
                $jsonConfig = Get-Content -LiteralPath $availableJsonConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
                
                # Update default settings from JSON
                foreach ($key in $jsonConfig.PSObject.Properties.Name) {
                    if ($script:DefaultSettings.ContainsKey($key)) {
                        $script:DefaultSettings[$key] = $jsonConfig.$key
                    }
                }
                
                # Check for encrypted password
                if ($jsonConfig.PSObject.Properties.Name -contains "SQLPasswordEncrypted" -and 
                    $jsonConfig.SQLPasswordEncrypted -eq $true -and 
                    -not [string]::IsNullOrEmpty($script:DefaultSettings.SQLPassword)) {
                    try {
                        $script:DefaultSettings.SQLPassword = Unprotect-Password -EncryptedPassword $script:DefaultSettings.SQLPassword
                    }
                    catch {
                        Write-ExtendedLog "Password decryption failed" -Level "WARNING"
                    }
                }
                
                Write-ExtendedLog "Configuration loaded from $jsonConfigPath" -Level "INFO" -Component "ConfigManager"
                return $true
            }
            catch {
                Write-ExtendedLog "Failed to load configuration: $($_.Exception.Message)" -Level "WARNING" -Component "ConfigManager"
                return $false
            }
        }
        else {
            Write-ExtendedLog "No configuration file found" -Level "INFO" -Component "ConfigManager"
            return $false
        }
    } 
    catch { 
        Write-ExtendedLog "Load settings failed: $($_.Exception.Message)" -Level "ERROR" -Component "ConfigManager"
        return $false
    }
}

function Add-GuiLogEntry {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Component = "VMDeletionTool"
    )

    if ($null -eq $script:txtLogs) { return }

    try {
        $entry = "{0} [{1}] [{2}] {3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level.ToUpper(), $Component, $Message
        if ($script:txtLogs.InvokeRequired) {
            [void]$script:txtLogs.BeginInvoke([Action[string]]{
                param($line)
                $script:txtLogs.AppendText($line + [Environment]::NewLine + [Environment]::NewLine)
                $script:txtLogs.SelectionStart = $script:txtLogs.TextLength
                $script:txtLogs.ScrollToCaret()
            }, $entry)
        }
        else {
            $script:txtLogs.AppendText($entry + [Environment]::NewLine + [Environment]::NewLine)
            $script:txtLogs.SelectionStart = $script:txtLogs.TextLength
            $script:txtLogs.ScrollToCaret()
        }
    }
    catch { }
}

function Show-GuiErrorNotice {
    if ($script:IsShowingLogErrorNotice) { return }
    if ($null -eq $form -or -not $form.Visible) { return }

    try {
        $script:IsShowingLogErrorNotice = $true
        [System.Windows.Forms.MessageBox]::Show(
            "ERROR: Go to Logs to see it.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
    catch { }
    finally {
        $script:IsShowingLogErrorNotice = $false
    }
}
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    Write-ExtendedLog -Message $Message -Level $Level
}

function Write-ExtendedLog {
    param(
        [string]$Message,
        [string]$Level = "INFO",
        [string]$Component = "VMDeletionTool",
        [switch]$SendEmail
    )
    
    # Log Directory Management
    $logDir = $script:DefaultSettings.LogPath
    $maxLogFiles = 10
    
    # Ensure log directory exists
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory | Out-Null
    }
    
    # Log Rotation
    $logFiles = Get-ChildItem "$logDir\VMDeletion_*.log" -ErrorAction SilentlyContinue |
    Sort-Object CreationTime -Descending
    
    if ($logFiles.Count -ge $maxLogFiles) {
        $oldestLog = $logFiles[-1].FullName
        Remove-Item $oldestLog -ErrorAction SilentlyContinue
    }
    
    # Create Log Entry
    $logEntry = [PSCustomObject]@{
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Level = $Level.ToUpper()
        Message = $Message
        Component = $Component
        User = "$env:USERDOMAIN\$env:USERNAME"
        Machine = $env:COMPUTERNAME
    }
    
    # Write to log file
    $logFile = Join-Path $logDir "VMDeletion_$(Get-Date -Format 'yyyyMMdd').log"
    
    # Serialize to JSON
    $jsonLogEntry = ($logEntry | ConvertTo-Json -Compress) + "`r`n"
    $jsonLogEntry | Out-File -FilePath $logFile -Append -Encoding UTF8
    
    # Console Output
    $formattedMessage = "{0} [{1}] [{2}] {3} - {4}" -f 
        $logEntry.Timestamp, 
        $logEntry.Level, 
        $logEntry.Component, 
        $logEntry.User, 
        $logEntry.Message
    if ($NoGui -or $Help) {
        $consoleColor = switch ($logEntry.Level) {
            "ERROR"   { "Red" }
            "WARNING" { "Yellow" }
            "INFO"    { "Gray" }
            default   { "White" }
        }
        Write-Host $formattedMessage -ForegroundColor $consoleColor
    }
    
    if ($Level -eq "ERROR") {
        Add-GuiLogEntry -Message $Message -Level $Level -Component $Component
        Show-GuiErrorNotice
    }

    # Email Notification for Critical Errors
    if ($SendEmail -and $Level -eq "ERROR") {
        Send-ErrorNotificationEmail -LogEntry $logEntry
    }
}

function Send-ErrorNotificationEmail {
    param($LogEntry)
    
    try {
        $emailParams = @{
            From = $script:DefaultSettings.EmailNotification
            To = $script:DefaultSettings.EmailNotification
            Subject = "VM Deletion Tool - Critical Error"
            Body = @"
<html>
<body>
<h2>Critical Error Notification</h2>
<table border='1' cellpadding='5'>
<tr><th>Timestamp</th><td>$($LogEntry.Timestamp)</td></tr>
<tr><th>Level</th><td>$($LogEntry.Level)</td></tr>
<tr><th>Message</th><td>$($LogEntry.Message)</td></tr>
<tr><th>User</th><td>$($LogEntry.User)</td></tr>
<tr><th>Machine</th><td>$($LogEntry.Machine)</td></tr>
</table>
</body>
</html>
"@
            BodyAsHtml = $true
            SmtpServer = $script:DefaultSettings.SMTPServer
        }
        
        Send-MailMessage @emailParams
    }
    catch {
        # Fallback logging if email fails
        $fallbackLogPath = Join-Path $script:DefaultSettings.LogPath "email_errors.log"
        "Failed to send email: $($_.Exception.Message)" | 
            Out-File -FilePath $fallbackLogPath -Append
    }
}

function Save-EmailGroups {
    # שמירת קבוצות מה-ListBoxes להגדרות
    if ($script:lstGroupsCommon) {
        $groups = @()
        for ($i = 0; $i -lt $script:lstGroupsCommon.Items.Count; $i++) {
            $groups += $script:lstGroupsCommon.Items[$i].ToString()
        }
        $script:DefaultSettings.GroupsCommon = $groups -join ","
    }
    
    if ($script:lstGroupsWindows) {
        $groups = @()
        for ($i = 0; $i -lt $script:lstGroupsWindows.Items.Count; $i++) {
            $groups += $script:lstGroupsWindows.Items[$i].ToString()
        }
        $script:DefaultSettings.GroupsWindows = $groups -join ","
    }
    
    if ($script:lstGroupsLinux) {
        $groups = @()
        for ($i = 0; $i -lt $script:lstGroupsLinux.Items.Count; $i++) {
            $groups += $script:lstGroupsLinux.Items[$i].ToString()
        }
        $script:DefaultSettings.GroupsLinux = $groups -join ","
    }
}


#endregion

#region SQL Functions

function Save-VMDeletionToArchive {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$VMDetails,
        [string]$Reason,
        [bool]$SCCMDeleted,
        [bool]$ADDeleted,
        [bool]$DNSDeleted,
        [bool]$VMDeleted,
        [string]$Status
    )

    try {
        # ✅ הוסף בדיקות בהתחלה
        Write-Log "💾 SAVE-VMDELETIONTOARCHIVE CALLED" -Level "INFO"
        Write-Log "📋 Current Settings: Server='$($script:DefaultSettings.SQLServer)', DB='$($script:DefaultSettings.SQLDatabase)', Table='$($script:DefaultSettings.SQLTable)'" -Level "INFO"
        
        # Build tracking data with validation for NOT NULL fields
        $trackingData = @{
            # NOT NULL fields - must have values
            Timestamp  = Get-Date
            ServerName = if ([string]::IsNullOrWhiteSpace($VMDetails.ServerName)) { 
                "UNKNOWN" 
            } else { 
                $VMDetails.ServerName 
            }
            Action     = "Delete"
            Initiator  = "$env:USERDOMAIN\$env:USERNAME"
            Reason     = if ([string]::IsNullOrWhiteSpace($Reason)) { 
                "No reason provided" 
            } else { 
                $Reason 
            }
            Platform   = if ([string]::IsNullOrWhiteSpace($VMDetails.Platform)) { 
                "Virtual Server" 
            } else { 
                $VMDetails.Platform 
            }
            Network    = if ([string]::IsNullOrWhiteSpace($VMDetails.Network)) { 
                "Unknown" 
            } else { 
                $VMDetails.Network 
            }
            vCenter    = if ([string]::IsNullOrWhiteSpace($global:DefaultVIServer.Name)) { 
                "Unknown" 
            } else { 
                $global:DefaultVIServer.Name 
            }
            
            # NULLABLE fields - can be NULL
            Owner           = $VMDetails.Owner
            Cluster         = $VMDetails.Cluster
            OperatingSystem = $VMDetails.OperatingSystem
            PowerState      = $VMDetails.PowerState
            PowerStateDate  = $VMDetails.PowerStateDate
            CPUCores        = $VMDetails.CPUCores
            MemoryGB        = $VMDetails.MemoryGB
            TotalDisksGB    = $VMDetails.TotalDisksGB
            NBLastBackup    = $VMDetails.NBLastBackup
        }

        $connectionString = Get-SQLConnectionString
        
        # ✅ בדוק חיבור לפני כתיבה
        Write-Log "🔗 Archive Connection: $($connectionString -replace 'Password=[^;]*', 'Password=***')" -Level "INFO"
        
        $testConn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $testConn.Open()
        $preCmd = $testConn.CreateCommand()
        $preCmd.CommandText = "SELECT @@SERVERNAME as Srv, DB_NAME() as DB, COUNT(*) as Cnt FROM $($script:DefaultSettings.SQLTable)"
        $preReader = $preCmd.ExecuteReader()
        if ($preReader.Read()) {
            Write-Log "💾 BEFORE INSERT: Server='$($preReader['Srv'])', DB='$($preReader['DB'])', Records=$($preReader['Cnt'])" -Level "INFO"
        }
        $preReader.Close()
        $testConn.Close()

        if ([string]::IsNullOrWhiteSpace($connectionString)) {
            throw "SQL Connection string is empty"
        }

        # Create table if needed
        #New-VMTrackingTable -ConnectionString $connectionString

        # Write record
        $writeResult = Write-VMTrackingRecord -ConnectionString $connectionString -Data $trackingData
        
        # ✅ בדוק אחרי כתיבה
        $testConn2 = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $testConn2.Open()
        $postCmd = $testConn2.CreateCommand()
        $postCmd.CommandText = "SELECT COUNT(*) FROM $($script:DefaultSettings.SQLTable)"
        $postCount = $postCmd.ExecuteScalar()
        Write-Log "💾 AFTER INSERT: Records=$postCount" -Level "INFO"
        $testConn2.Close()
        
        if ($writeResult) {
            return @{
                Success = $true
                Message = "Archived VM deletion info for $($VMDetails.ServerName)"
            }
        }
        else {
            throw "Write-VMTrackingRecord returned false"
        }
    }
    catch {
        Write-Log "❌ Save-VMDeletionToArchive failed: $($_.Exception.Message)" -Level "ERROR"
        return @{
            Success = $false
            Message = "Failed to archive: $($_.Exception.Message)"
        }
    }
}

function New-VMTrackingTable {
    param (
        [string]$ConnectionString
    )

    $query = @"
IF NOT EXISTS (SELECT * FROM $($script:DefaultSettings.SQLDatabase) WHERE name = '$($script:DefaultSettings.SQLTable)')
BEGIN
    CREATE TABLE $($script:DefaultSettings.SQLTable) (
        Timestamp       DATETIME2       NULL,
        ServerName      NVARCHAR(128)   NULL,
        Action          NVARCHAR(50)    NULL,
        Owner           NVARCHAR(128)   NULL,
        Initiator       NVARCHAR(128)   NULL,
        Reason          NVARCHAR(MAX)   NULL,
        Platform        NVARCHAR(50)    NULL,
        Network         NVARCHAR(50)    NULL,
        vCenter         NVARCHAR(128)   NULL,
        Cluster         NVARCHAR(128)   NULL,
        OperatingSystem NVARCHAR(128)   NULL,
        PowerState      NVARCHAR(50)    NULL,
        PowerStateDate  DATETIME2       NULL,
        CPUCores        INT             NULL,
        MemoryGB        FLOAT           NULL,
        TotalDisksGB    FLOAT           NULL,
        NBLastBackup    NVARCHAR(128)   NULL
    )
END
"@

    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $connection.Open()
        $command.ExecuteNonQuery()
        $connection.Close()
        Write-Log "$script:DefaultSettings.SQLTable table verified/created"
    }
    catch {
        Write-Log "Failed to create $script:DefaultSettings.SQLTable table: $($_.Exception.Message)" -Level "ERROR"
    }
}

function Write-VMTrackingRecord {
    param (
        [string]$ConnectionString,
        [hashtable]$Data
    )

    # ✅ Validate and ensure NOT NULL fields have values
    $notNullFields = @{
        'Timestamp'  = { Get-Date }
        'ServerName' = { "UNKNOWN" }
        'Action'     = { "Delete" }
        'Initiator'  = { "$env:USERDOMAIN\$env:USERNAME" }
        'Reason'     = { "No reason provided" }
        'Platform'   = { "Unknown" }
        'Network'    = { "Unknown" }
        'vCenter'    = { "Unknown" }
    }

    foreach ($field in $notNullFields.Keys) {
        if ([string]::IsNullOrWhiteSpace($Data[$field])) {
            $Data[$field] = & $notNullFields[$field]
            Write-Log "Field '$field' was empty, set to default: '$($Data[$field])'" -Level "WARNING"
        }
    }

    if ([string]::IsNullOrWhiteSpace($Data.ServerName)) {
        Write-Log "CRITICAL: Cannot insert record - ServerName is still empty after defaults" -Level "ERROR"
        return $false
    }

    Write-Log "=== Attempting SQL Insert ===" -Level "INFO"
    Write-Log "ServerName: $($Data.ServerName)" -Level "INFO"

    # Sanitize string fields
    $stringFields = @('ServerName', 'Action', 'Owner', 'Initiator', 'Reason', 
                      'Platform', 'Network', 'vCenter', 'Cluster', 'OperatingSystem', 
                      'PowerState', 'NBLastBackup')

    foreach ($field in $stringFields) {
        if ($Data.ContainsKey($field)) {
            if ([string]::IsNullOrWhiteSpace($Data[$field])) {
                if (-not $notNullFields.ContainsKey($field)) {
                    $Data[$field] = $null
                }
            } else {
                $Data[$field] = $Data[$field].ToString().Trim()
            }
        }
    }

    # Sanitize DateTime fields
    if ($Data.ContainsKey('Timestamp')) {
        if ($Data.Timestamp -isnot [datetime]) {
            try {
                $Data.Timestamp = [datetime]::Parse($Data.Timestamp)
            } catch {
                $Data.Timestamp = Get-Date
                Write-Log "Invalid Timestamp, using current time" -Level "WARNING"
            }
        }
    }

    if ($Data.ContainsKey('PowerStateDate') -and $Data.PowerStateDate) {
        if ($Data.PowerStateDate -isnot [datetime]) {
            try {
                $Data.PowerStateDate = [datetime]::Parse($Data.PowerStateDate)
            } catch {
                $Data.PowerStateDate = $null
                Write-Log "Invalid PowerStateDate, setting to NULL" -Level "WARNING"
            }
        }
    } else {
        $Data.PowerStateDate = $null
    }

    # Sanitize numeric fields
    $numericFields = @('CPUCores', 'MemoryGB', 'TotalDisksGB')

    foreach ($field in $numericFields) {
        if ($Data.ContainsKey($field)) {
            if ($null -eq $Data[$field] -or [string]::IsNullOrWhiteSpace($Data[$field].ToString())) {
                $Data[$field] = $null
            }
            else {
                try {
                    if ($field -eq 'CPUCores') {
                        $Data[$field] = [int]$Data[$field]
                    } else {
                        $Data[$field] = [decimal]$Data[$field]
                    }
                } 
                catch {
                    $Data[$field] = $null
                    Write-Log "Invalid $field format, setting to NULL" -Level "WARNING"
                }
            }
        }
    }

    # ✅ Transaction מפורש עם COMMIT
    $query = @"
BEGIN TRANSACTION;
SET IMPLICIT_TRANSACTIONS OFF;
INSERT INTO $($script:DefaultSettings.SQLTable) (
    TIMESTAMP, ServerName, Action, Owner, Initiator, Reason,
    Platform, Network, vCenter, Cluster, OperatingSystem,
    PowerState, PowerStateDate, CPUCores, MemoryGB, TotalDisksGB, NBLastBackup
)
VALUES (
    @Timestamp, @ServerName, @Action, @Owner, @Initiator, @Reason,
    @Platform, @Network, @vCenter, @Cluster, @OperatingSystem,
    @PowerState, @PowerStateDate, @CPUCores, @MemoryGB, @TotalDisksGB, @NBLastBackup
);
COMMIT TRANSACTION;
"@

    $connection = $null
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
        $connection.Open()
        Write-Log "SQL Connection opened successfully" -Level "INFO"
        
        # ✅ הדפס פרטי חיבור מדויקים
        $serverInfoCommand = $connection.CreateCommand()
        $serverInfoCommand.CommandText = "SELECT @@SERVERNAME as ServerName, DB_NAME() as DatabaseName, SYSTEM_USER as SystemUser"
        $infoReader = $serverInfoCommand.ExecuteReader()
        if ($infoReader.Read()) {
            $actualServer = $infoReader["ServerName"]
            $actualDB = $infoReader["DatabaseName"]
            $actualUser = $infoReader["SystemUser"]
            Write-Log "🎯 CONNECTED TO: Server='$actualServer', Database='$actualDB', User='$actualUser'" -Level "INFO"
        }
        $infoReader.Close()
        
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $command.CommandTimeout = 60

        Write-Log "Adding SQL parameters:" -Level "INFO"
        foreach ($key in $Data.Keys) {
            $value = if ($null -ne $Data[$key]) { $Data[$key] } else { [DBNull]::Value }
            $command.Parameters.AddWithValue("@$key", $value) | Out-Null
            $valueDisplay = if ($value -is [DBNull]) { "NULL" } else { $value }
            Write-Log "  @$key = $valueDisplay" -Level "INFO"
        }
        
        $rowsAffected = $command.ExecuteNonQuery()
        Write-Log "SQL ExecuteNonQuery completed. Rows affected: $rowsAffected" -Level "INFO"

        # ✅ בדיקת אימות מיידית - כמה רשומות יש עכשיו
        $countCommand = $connection.CreateCommand()
        $countCommand.CommandText = "SELECT COUNT(*) FROM $($script:DefaultSettings.SQLTable)"
        $totalCount = $countCommand.ExecuteScalar()
        Write-Log "📊 Total records in $($script:DefaultSettings.SQLTable) table: $totalCount" -Level "INFO"

        # ✅ בדיקה ספציפית לשרת הזה
        $verifyCommand = $connection.CreateCommand()
        $verifyCommand.CommandText = @"
SELECT COUNT(*) FROM $($script:DefaultSettings.SQLTable)
WHERE ServerName = @ServerName 
AND TIMESTAMP >= @MinTime
"@
        $verifyCommand.Parameters.AddWithValue("@ServerName", $Data.ServerName) | Out-Null
        $verifyCommand.Parameters.AddWithValue("@MinTime", (Get-Date).AddMinutes(-5)) | Out-Null
        
        $verifyCount = $verifyCommand.ExecuteScalar()
        Write-Log "🔍 Verification: Found $verifyCount record(s) for '$($Data.ServerName)' in last 5 minutes" -Level "INFO"

        # ✅ הדפס את הרשומה האחרונה בטבלה
        $lastRecordCommand = $connection.CreateCommand()
        $lastRecordCommand.CommandText = @"
SELECT TOP 1 
    TIMESTAMP,
    ServerName, 
    Action,
    Initiator,
    Reason
FROM $($script:DefaultSettings.SQLTable)
ORDER BY TIMESTAMP DESC
"@
        $lastReader = $lastRecordCommand.ExecuteReader()
        if ($lastReader.Read()) {
            Write-Log "📝 LAST RECORD IN TABLE: Time='$($lastReader["TIMESTAMP"])', Server='$($lastReader["ServerName"])', Action='$($lastReader["Action"])', User='$($lastReader["Initiator"])', Reason='$($lastReader["Reason"])'" -Level "INFO"
        } else {
            Write-Log "⚠️ No records found in table!" -Level "WARNING"
        }
        $lastReader.Close()

        # ✅ הדפס את הרשומה האחרונה לשרת הספציפי הזה
        $serverLastCommand = $connection.CreateCommand()
        $serverLastCommand.CommandText = @"
SELECT TOP 1 
    TIMESTAMP,
    ServerName,
    Action,
    Initiator
FROM $($script:DefaultSettings.SQLTable)
WHERE ServerName = @ServerName
ORDER BY TIMESTAMP DESC
"@
        $serverLastCommand.Parameters.AddWithValue("@ServerName", $Data.ServerName) | Out-Null
        $serverReader = $serverLastCommand.ExecuteReader()
        if ($serverReader.Read()) {
            Write-Log "📝 LAST RECORD FOR '$($Data.ServerName)': Time='$($serverReader["TIMESTAMP"])', Action='$($serverReader["Action"])', User='$($serverReader["Initiator"])'" -Level "INFO"
        } else {
            Write-Log "⚠️ No previous records found for server '$($Data.ServerName)'" -Level "INFO"
        }
        $serverReader.Close()

        if ($rowsAffected -gt 0) {
            Write-Log "✅ Successfully inserted $($script:DefaultSettings.SQLTable) record for $($Data.ServerName)" -Level "INFO"
            return $true
        } else {
            Write-Log "⚠️ SQL executed but 0 rows affected!" -Level "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "❌ Failed to insert tracking record: $($_.Exception.Message)" -Level "ERROR"
        Write-Log "Error Type: $($_.Exception.GetType().FullName)" -Level "ERROR"
        if ($_.Exception.InnerException) {
            Write-Log "Inner Exception: $($_.Exception.InnerException.Message)" -Level "ERROR"
        }
        return $false
    }
    finally {
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
            Write-Log "Connection closed" -Level "INFO"
        }
    }
}

function Get-SQLConnectionString {
    $server = $script:DefaultSettings.SQLServer
    $database = $script:DefaultSettings.SQLDatabase
    
    if ([string]::IsNullOrWhiteSpace($server)) {
        Write-Log "SQL Server not configured" -Level "ERROR"
        return $null
    }
    
    if ($script:DefaultSettings.SQLIntegratedAuth) {
        # Windows Authentication - משתמש בחשבון Windows הנוכחי
        $connectionString = "Server=$server;Database=$database;Integrated Security=True;Connection Timeout=3;"
        Write-Log "Using Windows Authentication for SQL"
    } 
    else {
        # SQL Server Authentication - משתמש ביוזר וסיסמא
        $user = $script:DefaultSettings.SQLUsername
        $pass = $script:DefaultSettings.SQLPassword
        
        if ([string]::IsNullOrWhiteSpace($user)) {
            Write-Log "SQL Username not configured" -Level "ERROR"
            return $null
        }
        
        $connectionString = "Server=$server;Database=$database;User Id=$user;Password=$pass;Connection Timeout=3;"
        Write-Log "Using SQL Authentication for SQL (User: $user)"
    }
    
    return $connectionString
}

function Test-SQLConnection {
    param([string]$ConnectionString)
    
    $result = @{
        Success = $false
        Message = ""
    }
    
    if ([string]::IsNullOrWhiteSpace($ConnectionString)) {
        $result.Message = "Connection string is empty"
        return $result
    }
    
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $ConnectionString
        $connection.Open()
        
        # בדיקה שאפשר לקרוא מה-DB
        $command = New-Object System.Data.SqlClient.SqlCommand("SELECT 1", $connection)
        $command.ExecuteScalar() | Out-Null
        
        $connection.Close()
        $result.Success = $true
        $result.Message = "Connected successfully to $($script:DefaultSettings.SQLServer)"
    }
    catch [System.Data.SqlClient.SqlException] {
        $sqlError = $_.Exception
        switch ($sqlError.Number) {
            18456 { $result.Message = "Login failed - check username/password" }
            4060  { $result.Message = "Database '$($script:DefaultSettings.SQLDatabase)' not found" }
            53    { $result.Message = "Server not found - check server name" }
            -1    { $result.Message = "Connection timeout - server unreachable" }
            default { $result.Message = "SQL Error $($sqlError.Number): $($sqlError.Message)" }
        }
    }
    catch {
        $result.Message = "Connection failed: $($_.Exception.Message)"
    }
    
    return $result
}

function Test-SQLTableExists {
    param(
        [string]$ConnectionString,
        [string]$TableName
    )
    
    $result = @{
        Success = $false
        Exists  = $false
        Message = ""
        Columns = @()
    }
    
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $ConnectionString
        $connection.Open()
        
        $query = @"
SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = @TableName
ORDER BY ORDINAL_POSITION
"@
        
        $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
        $command.Parameters.AddWithValue("@TableName", $TableName) | Out-Null
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $dataTable = New-Object System.Data.DataTable
        $adapter.Fill($dataTable) | Out-Null
        
        $connection.Close()
        
        $result.Success = $true
        
        if ($dataTable.Rows.Count -gt 0) {
            $result.Exists = $true
            $result.Columns = $dataTable.Rows | ForEach-Object { $_.COLUMN_NAME }
            $result.Message = "Table exists with $($dataTable.Rows.Count) columns"
        } else {
            $result.Exists = $false
            $result.Message = "Table does not exist"
        }
    }
    catch {
        $result.Message = "Error checking table: $($_.Exception.Message)"
    }
    
    return $result
}

function Get-VMDetailsForArchive {
    param([string]$VMName)
    
    $details = @{
        # === שדות למייל (Send-VMDeletionNotification) ===
        VMName          = $VMName
        VCenterName     = $null
        ClusterName     = $null
        FolderName      = $null
        ServerOwner     = $null
        TotalCPUs       = $null
        TotalMemoryGB   = $null
        TotalDisksGB    = $null
        OperatingSystem = $null
        PowerState      = $null
        
        # === שדות ל-Archive (Save-VMDeletionToArchive) ===
        Timestamp       = Get-Date
        ServerName      = $VMName
        Action          = "Delete"
        Owner           = $null
        Initiator       = "$env:USERDOMAIN\$env:USERNAME"
        Reason          = $null
        Platform        = "Virtual Server"
        Network         = "External"
        vCenter         = $null
        Cluster         = $null
        PowerStateDate  = $null
        CPUCores        = $null
        MemoryGB        = $null
        NBLastBackup    = $null
    }
    
    try {
        $vm = Get-VM -Name $VMName -ErrorAction Stop
       
        # vCenter Name
        if ($global:DefaultVIServer -and $global:DefaultVIServer.Name) {
            $details.VCenterName = $global:DefaultVIServer.Name
            $details.vCenter = $global:DefaultVIServer.Name  # לArchive
        }
        
        # Cluster Name
        $cluster = Get-Cluster -VM $vm -ErrorAction SilentlyContinue
        if ($cluster) { 
            $details.ClusterName = $cluster.Name
            $details.Cluster = $cluster.Name  # לArchive
        }

        # Folder Name
        $folder = $vm.Folder
        if ($folder) {
            $details.FolderName = $folder.Name
        }

        # Server Owner from Annotation (srvOWN1)
        try {
            $annotation = Get-Annotation -Entity $vm -CustomAttribute "srvOWN1" -ErrorAction SilentlyContinue
            if ($annotation -and $annotation.Value) {
                $details.ServerOwner = $annotation.Value
                $details.Owner = $annotation.Value  # לArchive
            }
        }
        catch {
            # Try from Notes
            if ($vm.Notes -match "srvOWN1[:\s]*(\S+)") {
                $details.ServerOwner = $matches[1]
                $details.Owner = $matches[1]
            }
        }
        
        # NB_LAST_BACKUP
        try {
            $nbAttr = Get-Annotation -Entity $vm -CustomAttribute "NB_LAST_BACKUP" -ErrorAction SilentlyContinue
            if ($nbAttr -and $nbAttr.Value) {
                $details.NBLastBackup = $nbAttr.Value
            }
        }
        catch { 
            $details.NBLastBackup = $null 
        }

        # srvPSDate
        try {
            $psDate = Get-Annotation -Entity $vm -CustomAttribute "srvPSDate" -ErrorAction SilentlyContinue
            if ($psDate -and $psDate.Value) {
                try {
                    $details.PowerStateDate = [datetime]::Parse($psDate.Value)
                }
                catch {
                    $details.PowerStateDate = $null
                }
            }
        } 
        catch { 
            $details.PowerStateDate = $null 
        }

        # Power State
        if ($vm.PowerState) {
            $details.PowerState = $vm.PowerState.ToString()
        }
        
        # CPUs
        if ($vm.NumCpu) {
            $details.TotalCPUs = [int]$vm.NumCpu
            $details.CPUCores = [int]$vm.NumCpu  # לArchive
        }
        
        # Memory (GB)
        if ($vm.MemoryGB) {
            $details.TotalMemoryGB = [math]::Round($vm.MemoryGB, 2)
            $details.MemoryGB = [math]::Round($vm.MemoryGB, 2)  # לArchive
        }
        
        # Total Disk Size (GB)
        $disks = Get-HardDisk -VM $vm -ErrorAction SilentlyContinue
        if ($disks) {
            $totalDiskGB = ($disks | Measure-Object -Property CapacityGB -Sum).Sum
            if ($totalDiskGB) {
                $details.TotalDisksGB = [math]::Round($totalDiskGB, 2)
            }
        }
        
        # Operating System
        $osFullName = $vm.Guest.OSFullName
        $guestId = $vm.GuestId
        
        if ($osFullName) { 
            $details.OperatingSystem = $osFullName 
        } 
        elseif ($guestId) { 
            $details.OperatingSystem = $guestId 
        }
    }
    catch {
        Write-Log "Error getting VM details for $VMName : $($_.Exception.Message)" -Level "WARNING"
    }
    
    return $details
}

function Get-ArchiveData {

    param([string]$ConnectionString, [string]$TableName)
    
    $result = @{
        Success = $false
        Message = ""
        Data    = $null
    }
    
    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $ConnectionString
        $connection.Open()
        
        $query = "SELECT * FROM [$TableName] ORDER BY [TIMESTAMP] DESC"
        $command = New-Object System.Data.SqlClient.SqlCommand($query, $connection)
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter($command)
        $dataTable = New-Object System.Data.DataTable
        $adapter.Fill($dataTable) | Out-Null
        
        $connection.Close()
        
        $result.Success = $true
        $result.Data = $dataTable
        $result.Message = "Retrieved $($dataTable.Rows.Count) records"
    }
    catch {
        $result.Message = "Error retrieving data: $($_.Exception.Message)"
    }
    
    return $result
}

function Export-ArchiveToFile {
    param(
        [System.Data.DataTable]$Data,
        [string]$FilePath,
        [string]$Format  # "CSV" or "XLSX"
    )
    
    $result = @{
        Success = $false
        Message = ""
    }
    
    try {
        if ($Format -eq "CSV") {
            $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
            $result.Success = $true
            $result.Message = "Exported to CSV: $FilePath"
        }
        elseif ($Format -eq "XLSX") {
            # Check if ImportExcel module is available
            if (Get-Module -ListAvailable -Name ImportExcel) {
                Import-Module ImportExcel
                $Data | Export-Excel -Path $FilePath -AutoSize -TableName "VMDeletionArchive" -FreezeTopRow
                $result.Success = $true
                $result.Message = "Exported to Excel: $FilePath"
            }
            else {
                # Fallback to CSV if ImportExcel not available
                $csvPath = $FilePath -replace '\.xlsx$', '.csv'
                $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                $result.Success = $true
                $result.Message = "ImportExcel module not found. Exported to CSV: $csvPath"
            }
        }
    }
    catch {
        $result.Message = "Export error: $($_.Exception.Message)"
    }
    
    return $result
}

function Export-ArchiveToExcel {
    param (
        [Parameter(Mandatory = $true)]
        $Data,
        [Parameter(Mandatory = $true)]
        [string]$FilePath
    )
    
    try {
        # יצירת אובייקט Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # יצירת Workbook חדש
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "Archive Data"
        
        # קבלת העמודות
        $columns = $Data | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
        
        # כתיבת כותרות
        $col = 1
        foreach ($column in $columns) {
            $worksheet.Cells.Item(1, $col) = $column
            $worksheet.Cells.Item(1, $col).Font.Bold = $true
            $worksheet.Cells.Item(1, $col).Interior.ColorIndex = 15  # אפור
            $col++
        }
        
        # כתיבת הנתונים
        $row = 2
        foreach ($item in $Data) {
            $col = 1
            foreach ($column in $columns) {
                $worksheet.Cells.Item($row, $col) = $item.$column
                $col++
            }
            $row++
        }
        
        # התאמת רוחב עמודות
        $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        
        # שמירה
        $workbook.SaveAs($FilePath)
        $workbook.Close($true)
        $excel.Quit()
        
        # שחרור זיכרון
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return @{ Success = $true; Message = "Exported successfully to: $FilePath" }
    }
    catch {
        return @{ Success = $false; Message = "Export failed: $($_.Exception.Message)" }
    }
}

#endregion SQL Functions

#region Deletion Functions

function Remove-SCCMDevice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$SCCMServer,
        
        [Parameter(Mandatory = $true)]
        [string]$SiteCode
    )
    
    $result = @{
        Success = $false
        Message = ""
        Found   = $false
    }
    
    try {
        $namespace = "root\SMS\site_$SiteCode"
        
        $query = "SELECT * FROM SMS_R_System WHERE Name = '$ComputerName'"
        $device = Get-WmiObject -ComputerName $SCCMServer -Namespace $namespace -Query $query -ErrorAction Stop
        
        if ($null -eq $device) {
            $result.Message = "Device not found in SCCM"
            $result.Found = $false
            $result.Success = $true
            return $result
        }
        
        $result.Found = $true
        $deviceList = @($device)
        $deletedCount = 0
        
        foreach ($dev in $deviceList) {
            $resourceId = $dev.ResourceId
            $deviceToDelete = Get-WmiObject -ComputerName $SCCMServer -Namespace $namespace `
                -Query "SELECT * FROM SMS_R_System WHERE ResourceId = $resourceId" -ErrorAction Stop
            
            if ($null -ne $deviceToDelete) {
                $deviceToDelete.Delete()
                $deletedCount++
            }
        }
        
        $result.Success = $true
        $result.Message = "Deleted $deletedCount record(s) from SCCM"
    } 
    catch [System.UnauthorizedAccessException] {
        $result.Message = "Access denied to SCCM server"
    }
    catch [System.Runtime.InteropServices.COMException] {
        $result.Message = "Cannot connect to SCCM server"
    }
    catch {
        $result.Message = "SCCM Error: $($_.Exception.Message)"
    }
    
    return $result
}

function Remove-ADDevice {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [string]$DomainController,
        
        [Parameter(Mandatory = $false)]
        [string]$SearchBase 
    )

    $result = @{
        Success  = $false
        Message  = ""
        Found    = $false
        Location = ""
    }

    try {
        if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
            $result.Message = "ActiveDirectory module not installed"
            return $result
        }

        Import-Module ActiveDirectory -ErrorAction Stop

        # קבע את ה-SearchBase - פרמטר > הגדרות > ברירת מחדל
        $effectiveSearchBase = if (-not [string]::IsNullOrWhiteSpace($SearchBase)) {
            $SearchBase
        } elseif (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.ADSearchBase)) {
            $script:DefaultSettings.ADSearchBase
        } else {
            $null  # יחפש בכל ה-AD
        }

        # קבע את ה-DomainController
        $effectiveDC = if (-not [string]::IsNullOrWhiteSpace($DomainController)) {
            $DomainController
        } elseif (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.DomainController)) {
            $script:DefaultSettings.DomainController
        } else {
            $null
        }

        # בנה את הפרמטרים לחיפוש
        $searchParams = @{
            Filter      = "Name -eq '$ComputerName'"
            ErrorAction = 'Stop'
        }

        # הוסף SearchBase אם הוגדר
        if (-not [string]::IsNullOrWhiteSpace($effectiveSearchBase)) {
            $searchParams.SearchBase = $effectiveSearchBase
            $searchParams.SearchScope = 'Subtree'  # חפש גם בתתי-OU
            Write-Log "Searching for '$ComputerName' in OU: $effectiveSearchBase" -Level "INFO"
        } else {
            Write-Log "Searching for '$ComputerName' in entire AD" -Level "WARNING"
        }

        # הוסף DC אם הוגדר
        if (-not [string]::IsNullOrWhiteSpace($effectiveDC)) {
            $searchParams.Server = $effectiveDC
        }

        # חפש את המחשב
        $computer = $null
        try {
            $computer = Get-ADComputer @searchParams
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $result.Message = "Server not found in AD"
            $result.Found = $false
            $result.Success = $true
            return $result
        }

        if ($null -eq $computer) {
            $result.Message = "Server not found "
            $result.Found = $false
            $result.Success = $true
            return $result
        }

        # נמצא!
        $result.Found = $true
        $result.Location = $computer.DistinguishedName
        Write-Log "Found computer at: $($computer.DistinguishedName)" -Level "INFO"

        # מחק את האובייקט
        $removeParams = @{
            Identity    = $computer.DistinguishedName
            Confirm     = $false
            ErrorAction = 'Stop'
        }

        if (-not [string]::IsNullOrWhiteSpace($effectiveDC)) {
            $removeParams.Server = $effectiveDC
        }

        Remove-ADComputer @removeParams

        $result.Success = $true
        $result.Message = "Server object deleted successfully"
        Write-Log "✅ Deleted AD object: $ComputerName" -Level "INFO"

    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $result.Message = "Server not found in specified OU"
        $result.Found = $false
        $result.Success = $true
    }
    catch [System.UnauthorizedAccessException] {
        $result.Message = "Access denied - insufficient AD permissions"
        Write-Log "❌ AD Access denied for: $ComputerName" -Level "ERROR"
    }
    catch {
        $result.Message = "AD Error: $($_.Exception.Message)"
        Write-Log "❌ AD Error: $($_.Exception.Message)" -Level "ERROR"
    }

    return $result
}

function Remove-DNSRecords {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [string]$DNSServer,
        
        [Parameter(Mandatory = $true)]
        [string]$DNSZone
    )
    
    $result = @{
        Success      = $false
        Message      = ""
        Found        = $false
        DeletedCount = 0
    }
    
    try {
        if (-not (Get-Module -ListAvailable -Name DnsServer)) {
            $result.Message = "DnsServer module not installed"
            return $result
        }
        
        Import-Module DnsServer -ErrorAction Stop
        
        $deletedRecords = 0
        $foundRecords = 0
        
        # A Records
        try {
            $aRecords = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                -ZoneName $DNSZone `
                -Name $ComputerName `
                -RRType A `
                -ErrorAction SilentlyContinue
            
            if ($null -ne $aRecords) {
                $foundRecords += @($aRecords).Count
                
                foreach ($record in $aRecords) {
                    Remove-DnsServerResourceRecord -ComputerName $DNSServer `
                        -ZoneName $DNSZone `
                        -Name $ComputerName `
                        -RRType A `
                        -RecordData $record.RecordData.IPv4Address.IPAddressToString `
                        -Force `
                        -ErrorAction Stop
                    $deletedRecords++
                }
            }
        }
        catch [Microsoft.Management.Infrastructure.CimException] { }
        
        # PTR Records
        try {
            $aRecordsForPtr = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                -ZoneName $DNSZone `
                -Name $ComputerName `
                -RRType A `
                -ErrorAction SilentlyContinue
            
            if ($null -ne $aRecordsForPtr) {
                foreach ($aRecord in $aRecordsForPtr) {
                    $ipAddress = $aRecord.RecordData.IPv4Address.IPAddressToString
                    $ipParts = $ipAddress.Split('.')
                    $reverseZone = "$($ipParts[2]).$($ipParts[1]).$($ipParts[0]).in-addr.arpa"
                    $ptrName = $ipParts[3]
                    
                    try {
                        $ptrRecord = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                            -ZoneName $reverseZone `
                            -Name $ptrName `
                            -RRType PTR `
                            -ErrorAction SilentlyContinue
                        
                        if ($null -ne $ptrRecord) {
                            $foundRecords++
                            Remove-DnsServerResourceRecord -ComputerName $DNSServer `
                                -ZoneName $reverseZone `
                                -Name $ptrName `
                                -RRType PTR `
                                -Force `
                                -ErrorAction SilentlyContinue
                            $deletedRecords++
                        }
                    }
                    catch { }
                }
            }
        }
        catch { }
        
        # AAAA Records
        try {
            $aaaaRecords = Get-DnsServerResourceRecord -ComputerName $DNSServer `
                -ZoneName $DNSZone `
                -Name $ComputerName `
                -RRType AAAA `
                -ErrorAction SilentlyContinue
            
            if ($null -ne $aaaaRecords) {
                $foundRecords += @($aaaaRecords).Count
                
                Remove-DnsServerResourceRecord -ComputerName $DNSServer `
                    -ZoneName $DNSZone `
                    -Name $ComputerName `
                    -RRType AAAA `
                    -Force `
                    -ErrorAction SilentlyContinue
                $deletedRecords += @($aaaaRecords).Count
            }
        }
        catch { }
        
        $result.Found = ($foundRecords -gt 0)
        $result.DeletedCount = $deletedRecords
        
        if ($foundRecords -eq 0) {
            $result.Success = $true
            $result.Message = "No DNS records found"
        }
        elseif ($deletedRecords -gt 0) {
            $result.Success = $true
            $result.Message = "Deleted $deletedRecords DNS record(s)"
        }
        else {
            $result.Message = "Found records but failed to delete"
        }
    }
    catch [System.UnauthorizedAccessException] {
        $result.Message = "Access denied to DNS server"
    }
    catch {
        $result.Message = "DNS Error: $($_.Exception.Message)"
    }
    
    return $result
}

#endregion Deletion Functions

function Get-VCenterServerForEnvironment {
    param([string]$EnvironmentName)

    $envName = Resolve-VCenterEnvironment -Value $EnvironmentName
    foreach ($vc in $script:vCenterServers) {
        if ($vc.Name -eq $envName -or $vc.Server -eq $envName) {
            return @{ Name = $vc.Name; Server = $vc.Server }
        }
    }

    return @{ Name = $envName; Server = $envName }
}

function Import-VMDeletionCsv {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "Missing required -CsvPath.`nNext step: run with -NoGui -CsvPath 'C:\Input\vms.csv', or use -Help for examples."
    }

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "CSV file was not found: $Path`nNext step: verify the path exists and that the Task Scheduler account can read it."
    }

    try {
        $csvData = @(Import-Csv -LiteralPath $Path -ErrorAction Stop)
    }
    catch {
        throw "CSV could not be read: $Path`nReason: $($_.Exception.Message)`nNext step: check file encoding, delimiter, and read permissions."
    }

    if ($csvData.Count -eq 0) {
        throw "CSV file is empty: $Path`nNext step: add at least one row with a VMName value."
    }

    $hasVMName = $csvData | Get-Member -Name "VMName" -MemberType NoteProperty
    if (-not $hasVMName) {
        throw "CSV is missing required column: VMName.`nExpected minimum format:`nVMName`nSERVER01`nSERVER02`nNext step: rename the server-name column to VMName."
    }

    $rows = @()
    $skippedBlankRows = 0
    foreach ($row in $csvData) {
        $vmName = [string]$row.VMName
        if ([string]::IsNullOrWhiteSpace($vmName)) {
            $skippedBlankRows++
            continue
        }

        $rowEnvironment = $script:RunEnvironment
        if ($row.PSObject.Properties.Name -contains "Environment" -and -not [string]::IsNullOrWhiteSpace($row.Environment)) {
            $rowEnvironment = $row.Environment
        }
        elseif ($row.PSObject.Properties.Name -contains "vCenter" -and -not [string]::IsNullOrWhiteSpace($row.vCenter)) {
            $rowEnvironment = Resolve-VCenterEnvironment -Value $row.vCenter
        }

        $rows += [pscustomobject]@{
            VMName      = $vmName.Trim()
            Environment = $rowEnvironment
        }
    }

    if ($rows.Count -eq 0) {
        throw "CSV does not contain valid VMName values.`nNext step: fill the VMName column with at least one VM or Template name."
    }

    if ($skippedBlankRows -gt 0) {
        Write-Log "CSV validation skipped $skippedBlankRows blank VMName row(s)." -Level "WARNING"
    }

    Write-Log "CSV validation completed. Rows=$($rows.Count), Source='$Path'"
    return $rows
}

function Show-NoGuiUsage {
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    $usage = @"
VM Deletion Tool - NoGui usage

Required:
  -NoGui -CsvPath <path-to-csv>

Common example:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$scriptPath" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD

Safe validation example, no cleanup actions enabled:
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$scriptPath" -NoGui -CsvPath "C:\Input\vms.csv" -Environment PROD -SkipSCCM -SkipAD -SkipDNS -SkipVM -SkipArchive -SkipEmail

CSV minimum format:
  VMName
  SERVER01
  SERVER02

CSV recommended format:
  VMName,Environment
  SERVER01,PROD
  SERVER02,MGMT
  SERVER03,VDI

Exit codes:
  0  Completed without VM-level failures
  1  Startup validation failed or one/more VM failures occurred
"@
    Write-Host $usage -ForegroundColor White
}

function Get-NoGuiActionSummary {
    $enabled = @()
    $skipped = @()
    foreach ($item in @(
        @{ Name = "SCCM";    Skipped = [bool]$SkipSCCM },
        @{ Name = "AD";      Skipped = [bool]$SkipAD },
        @{ Name = "DNS";     Skipped = [bool]$SkipDNS },
        @{ Name = "VM";      Skipped = [bool]$SkipVM },
        @{ Name = "Archive"; Skipped = [bool]$SkipArchive },
        @{ Name = "Email";   Skipped = [bool]$SkipEmail }
    )) {
        if ($item.Skipped) { $skipped += $item.Name } else { $enabled += $item.Name }
    }

    return [pscustomobject]@{
        Enabled = if ($enabled.Count -gt 0) { $enabled -join ", " } else { "None" }
        Skipped = if ($skipped.Count -gt 0) { $skipped -join ", " } else { "None" }
    }
}
function Invoke-HeadlessCsvDeletion {
    param([string]$Path)

    $actions = Get-NoGuiActionSummary
    Write-Log "NoGui run started. CSV='$Path' Environment='$script:RunEnvironment' Enabled=[$($actions.Enabled)] Skipped=[$($actions.Skipped)]"
    $rows = Import-VMDeletionCsv -Path $Path
    $rowCount = @($rows).Count
    $environmentGroupCount = @($rows | Group-Object Environment).Count
    Write-Log "NoGui will process $rowCount VM(s) across $environmentGroupCount environment group(s)."
    $successCount = 0
    $failedCount = 0

    if (-not $SkipVM) {
        if (-not (Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {
            Import-Module VMware.VimAutomation.Core -ErrorAction Stop
        }
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
    }

    $sqlReady = $false
    if (-not $SkipArchive) {
        $connectionString = Get-SQLConnectionString
        if (-not [string]::IsNullOrWhiteSpace($connectionString)) {
            $connTest = Test-SQLConnection -ConnectionString $connectionString
            if ($connTest.Success) {
                $tableCheck = Test-SQLTableExists -ConnectionString $connectionString -TableName $script:DefaultSettings.SQLTable
                $sqlReady = ($tableCheck.Success -and $tableCheck.Exists)
            }
        }
    }

    foreach ($group in ($rows | Group-Object Environment)) {
        $vcInfo = Get-VCenterServerForEnvironment -EnvironmentName $group.Name
        $connectedForGroup = $false

        if (-not $SkipVM) {
            try {
                Write-Log "Connecting to vCenter '$($vcInfo.Name)' ($($vcInfo.Server))"
                Connect-VIServer -Server $vcInfo.Server -ErrorAction Stop | Out-Null
                $connectedForGroup = $true
            }
            catch {
                Write-Log "Failed to connect to vCenter '$($vcInfo.Server)': $($_.Exception.Message)" -Level "ERROR"
                $failedCount += $group.Group.Count
                continue
            }
        }

        foreach ($row in $group.Group) {
            $vmName = $row.VMName
            $vmEnvironment = $row.Environment
            $vmErrors = @()
            Write-Log "NoGui deletion started for '$vmName' in environment '$vmEnvironment'"

            if ($SkipVM) {
                $vmDetails = @{ ServerName = $vmName; VMName = $vmName }
            }
            else {
                try { $vmDetails = Get-VMDetailsForArchive -VMName $vmName }
                catch { $vmDetails = @{ ServerName = $vmName; VMName = $vmName } }
            }

            $sccmDeleted = $false
            $adDeleted = $false
            $dnsDeleted = $false
            $vmDeleted = $false

            if (-not $SkipSCCM) {
                $sccmResult = Remove-SCCMDevice -ComputerName $vmName -SCCMServer $script:DefaultSettings.SCCMServer -SiteCode $script:DefaultSettings.SCCMSiteCode
                $sccmDeleted = $sccmResult.Found
                if (-not $sccmResult.Success) { $vmErrors += "SCCM: $($sccmResult.Message)" }
            }

            if (-not $SkipAD) {
                $adSearchBase = Get-ADSearchBaseForEnvironment -EnvironmentName $vmEnvironment
                $adResult = Remove-ADDevice -ComputerName $vmName -DomainController $script:DefaultSettings.DomainController -SearchBase $adSearchBase
                $adDeleted = $adResult.Found
                if (-not $adResult.Success) { $vmErrors += "AD: $($adResult.Message)" }
            }

            if (-not $SkipDNS) {
                $dnsResult = Remove-DNSRecords -ComputerName $vmName -DNSServer $script:DefaultSettings.DNSServer -DNSZone $script:DefaultSettings.DNSZone
                $dnsDeleted = $dnsResult.Found
                if (-not $dnsResult.Success) { $vmErrors += "DNS: $($dnsResult.Message)" }
            }

            if (-not $SkipVM) {
                try {
                    $template = Get-Template -Name $vmName -ErrorAction SilentlyContinue
                    if ($template) {
                        if ($script:DefaultSettings.DeleteFromDisk) {
                            Remove-Template -Template $template -DeletePermanently -Confirm:$false -ErrorAction Stop
                        }
                        else {
                            Remove-Template -Template $template -Confirm:$false -ErrorAction Stop
                        }
                    }
                    else {
                        $vm = Get-VM -Name $vmName -ErrorAction Stop
                        if ($vm.PowerState -eq "PoweredOn") {
                            Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null
                            $timeout = $script:DefaultSettings.PowerOffTimeout
                            $elapsed = 0
                            while ((Get-VM -Name $vmName -ErrorAction SilentlyContinue).PowerState -eq "PoweredOn" -and $elapsed -lt $timeout) {
                                Start-Sleep -Seconds 2
                                $elapsed += 2
                            }
                        }

                        if ($script:DefaultSettings.DeleteFromDisk) {
                            Remove-VM -VM $vm -DeletePermanently -Confirm:$false -ErrorAction Stop
                        }
                        else {
                            Remove-VM -VM $vm -Confirm:$false -ErrorAction Stop
                        }
                    }
                    $vmDeleted = $true
                }
                catch {
                    $vmErrors += "VM: $($_.Exception.Message)"
                }
            }

            $status = if ($vmErrors.Count -eq 0) { "Success" } else { "Failed" }

            if (-not $SkipArchive -and $sqlReady) {
                Save-VMDeletionToArchive -VMDetails $vmDetails -Reason $Reason -SCCMDeleted $sccmDeleted -ADDeleted $adDeleted -DNSDeleted $dnsDeleted -VMDeleted $vmDeleted -Status $status | Out-Null
            }

            if (-not $SkipEmail -and $script:DefaultSettings.EnableEmailNotification) {
                Send-VMDeletionNotification -VMDetails $vmDetails -Reason $Reason -OSType "Unknown" -DeletionSuccess $vmDeleted | Out-Null
            }

            if ($vmErrors.Count -eq 0) {
                $successCount++
                Write-Log "NoGui deletion completed successfully for '$vmName'"
            }
            else {
                $failedCount++
                Write-Log "NoGui deletion failed for '$vmName': $($vmErrors -join '; ')" -Level "ERROR"
            }
        }

        if ($connectedForGroup) {
            Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue
        }
    }

    Write-Log "NoGui CSV deletion completed. Success=$successCount Failed=$failedCount"
    if ($failedCount -gt 0) { return 1 }
    return 0
}

# Settings are loaded by Initialize-Configuration above. Do not reload from TEMP here.
if ($Help) {
    Show-NoGuiUsage
    exit 0
}

if ($NoGui) {
    try {
        if ([string]::IsNullOrWhiteSpace($CsvPath)) {
            Show-NoGuiUsage
            throw "Missing required -CsvPath."
        }

        $exitCode = Invoke-HeadlessCsvDeletion -Path $CsvPath
        if ($exitCode -eq 0) {
            Write-Host "Completed successfully. ExitCode=0" -ForegroundColor Green
        }
        else {
            Write-Host "Completed with failures. ExitCode=$exitCode. Next step: review the VMDeletion log and failed VM entries." -ForegroundColor Yellow
        }
        exit $exitCode
    }
    catch {
        Write-Log "NoGui run failed: $($_.Exception.Message)" -Level "ERROR"
        Write-Host "Startup failed. ExitCode=1. Next step: fix the error above and rerun with -Help if needed." -ForegroundColor Red
        exit 1
    }
}
#endregion

#region Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "VM Deletion Tool"
$form.Size = New-Object System.Drawing.Size(680, 930)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(46, 46, 58)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.TopMost = $false

# Modern frontend palette: calm charcoal surfaces, clear accent, no pure white glare.
$uiAccent = [System.Drawing.Color]::FromArgb(72, 166, 212)
$uiAccent2 = [System.Drawing.Color]::FromArgb(45, 135, 183)
$uiDanger = [System.Drawing.Color]::FromArgb(221, 74, 80)
$uiWarning = [System.Drawing.Color]::FromArgb(164, 142, 84)
$uiSuccess = [System.Drawing.Color]::FromArgb(62, 178, 122)
$uiText = [System.Drawing.Color]::FromArgb(238, 240, 244)
$uiMuted = [System.Drawing.Color]::FromArgb(188, 192, 202)
$uiButtonText = [System.Drawing.Color]::FromArgb(248, 249, 252)
$uiPanel = [System.Drawing.Color]::FromArgb(54, 54, 68)
$uiPanel2 = [System.Drawing.Color]::FromArgb(62, 62, 78)
$uiInput = [System.Drawing.Color]::FromArgb(34, 34, 43)
$uiInput2 = [System.Drawing.Color]::FromArgb(39, 39, 49)
$uiBorder = [System.Drawing.Color]::FromArgb(83, 83, 101)
$uiBorderSoft = [System.Drawing.Color]::FromArgb(68, 68, 84)
$uiBlack = [System.Drawing.Color]::FromArgb(46, 46, 58)

function Set-ModernButton {
    param(
        [System.Windows.Forms.Button]$Button,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor,
        [int]$BorderSize = 1,
        [int]$Radius = 12
    )
    $hoverColor = [System.Windows.Forms.ControlPaint]::Light($BackColor, 0.10)
    $borderColor = [System.Windows.Forms.ControlPaint]::Light($BackColor, 0.18)
    $Button.BackColor = $BackColor
    $Button.ForeColor = $ForeColor
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = $BorderSize
    $Button.FlatAppearance.BorderColor = $borderColor
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.UseVisualStyleBackColor = $false
    $Button.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
    $Button.Region = New-RoundedRegion -Width $Button.Width -Height $Button.Height -Radius $Radius
    $Button.Tag = @{ Radius = $Radius; BackColor = $BackColor; HoverColor = $hoverColor; BorderColor = $borderColor }
    $Button.Add_MouseEnter({
        param($sender, $eventArgs)
        $sender.BackColor = $sender.Tag.HoverColor
    })
    $Button.Add_MouseLeave({
        param($sender, $eventArgs)
        $sender.BackColor = $sender.Tag.BackColor
    })
    $Button.Add_Resize({
        param($sender, $eventArgs)
        $sender.Region = New-RoundedRegion -Width $sender.Width -Height $sender.Height -Radius $sender.Tag.Radius
    })
    $Button.Add_Paint({
        param($sender, $eventArgs)
        if ($sender.Width -le 0 -or $sender.Height -le 0) { return }
        $eventArgs.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $linePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(38, 255, 255, 255), 1)
        $eventArgs.Graphics.DrawLine($linePen, 8, 1, ($sender.Width - 9), 1)
        $linePen.Dispose()
    })
}

function New-RoundedRegion {
    param([int]$Width, [int]$Height, [int]$Radius)
    $diameter = [Math]::Max(2, $Radius * 2)
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $diameter, $diameter, 180, 90)
    $path.AddArc(($Width - $diameter - 1), 0, $diameter, $diameter, 270, 90)
    $path.AddArc(($Width - $diameter - 1), ($Height - $diameter - 1), $diameter, $diameter, 0, 90)
    $path.AddArc(0, ($Height - $diameter - 1), $diameter, $diameter, 90, 90)
    $path.CloseFigure()
    return (New-Object System.Drawing.Region -ArgumentList $path)
}

function Set-ModernGroup {
    param([System.Windows.Forms.GroupBox]$Group)
    $Group.BackColor = $uiPanel
    $Group.ForeColor = $uiAccent
    $Group.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
}

function Set-ModernInput {
    param([System.Windows.Forms.Control]$Control)
    $Control.BackColor = $uiInput
    $Control.ForeColor = $uiText
    if ($Control -is [System.Windows.Forms.TextBox] -or $Control -is [System.Windows.Forms.ListBox]) {
        $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    }
    if ($Control -is [System.Windows.Forms.ComboBox]) {
        $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    }
}

function Set-ModernButtonIfPresent {
    param(
        [string]$Name,
        [System.Drawing.Color]$BackColor,
        [System.Drawing.Color]$ForeColor = $uiButtonText,
        [int]$BorderSize = 1,
        [int]$Radius = 10
    )

    $buttonVar = Get-Variable -Name $Name -ErrorAction SilentlyContinue
    if ($buttonVar -and $buttonVar.Value -is [System.Windows.Forms.Button]) {
        Set-ModernButton -Button $buttonVar.Value -BackColor $BackColor -ForeColor $ForeColor -BorderSize $BorderSize -Radius $Radius
    }
}

function Set-ButtonPaddingTree {
    param([System.Windows.Forms.Control]$Root)

    foreach ($control in $Root.Controls) {
        if ($control -is [System.Windows.Forms.Button]) {
            $control.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
        }

        if ($control.Controls.Count -gt 0) {
            Set-ButtonPaddingTree -Root $control
        }
    }
}

function Set-AccessibleControlIfPresent {
    param(
        [string]$Name,
        [string]$AccessibleName,
        [string]$AccessibleDescription = "",
        [Nullable[int]]$TabIndex = $null
    )

    $controlVar = Get-Variable -Name $Name -ErrorAction SilentlyContinue
    if ($controlVar -and $controlVar.Value -is [System.Windows.Forms.Control]) {
        $control = $controlVar.Value
        $control.AccessibleName = $AccessibleName
        if (-not [string]::IsNullOrWhiteSpace($AccessibleDescription)) {
            $control.AccessibleDescription = $AccessibleDescription
        }
        if ($null -ne $TabIndex) {
            $control.TabIndex = [int]$TabIndex
        }
    }
}

function Set-AccessibilityTree {
    param([System.Windows.Forms.Control]$Root)

    foreach ($control in $Root.Controls) {
        if ($control -is [System.Windows.Forms.Label]) {
            $control.TabStop = $false
        }
        elseif ($control -is [System.Windows.Forms.Button]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::PushButton
            if ([string]::IsNullOrWhiteSpace($control.AccessibleName)) {
                $control.AccessibleName = $control.Text
            }
            if ($control.FlatAppearance.BorderSize -lt 1) {
                $control.FlatAppearance.BorderSize = 1
            }
        }
        elseif ($control -is [System.Windows.Forms.CheckBox]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::CheckButton
            if ([string]::IsNullOrWhiteSpace($control.AccessibleName)) {
                $control.AccessibleName = $control.Text
            }
        }
        elseif ($control -is [System.Windows.Forms.TextBox]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::Text
        }
        elseif ($control -is [System.Windows.Forms.ComboBox]) {
            $control.AccessibleRole = [System.Windows.Forms.AccessibleRole]::DropList
        }
        elseif ($control -is [System.Windows.Forms.ListView] -or $control -is [System.Windows.Forms.DataGridView]) {
            if ([string]::IsNullOrWhiteSpace($control.AccessibleName)) {
                $control.AccessibleName = $control.Name
            }
        }

        if ($control.Controls.Count -gt 0) {
            Set-AccessibilityTree -Root $control
        }
    }
}

function Set-DarkControlTree {
    param([System.Windows.Forms.Control]$Root)

    foreach ($control in $Root.Controls) {
        if ($control -is [System.Windows.Forms.TabPage]) {
            $control.BackColor = $uiPanel
            $control.ForeColor = $uiText
        }
        elseif ($control -is [System.Windows.Forms.GroupBox]) {
            Set-ModernGroup -Group $control
        }
        elseif ($control -is [System.Windows.Forms.Label]) {
            $control.BackColor = [System.Drawing.Color]::Transparent
            $control.ForeColor = $uiText
        }
        elseif ($control -is [System.Windows.Forms.CheckBox]) {
            $control.BackColor = $uiPanel
            $control.ForeColor = $uiText
        }
        elseif ($control -is [System.Windows.Forms.TextBox] -or
                $control -is [System.Windows.Forms.ComboBox] -or
                $control -is [System.Windows.Forms.NumericUpDown] -or
                $control -is [System.Windows.Forms.ListBox]) {
            Set-ModernInput -Control $control
        }
        elseif ($control -is [System.Windows.Forms.ListView]) {
            $control.BackColor = $uiInput
            $control.ForeColor = $uiText
            $control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        }
        elseif ($control -is [System.Windows.Forms.DataGridView]) {
            $control.BackgroundColor = $uiInput
            $control.GridColor = $uiBorder
            $control.ForeColor = $uiText
            $control.DefaultCellStyle.BackColor = $uiInput
            $control.DefaultCellStyle.ForeColor = $uiText
            $control.ColumnHeadersDefaultCellStyle.BackColor = $uiPanel2
            $control.ColumnHeadersDefaultCellStyle.ForeColor = $uiText
            $control.DefaultCellStyle.SelectionBackColor = $uiAccent2
            $control.DefaultCellStyle.SelectionForeColor = $uiButtonText
            $control.AlternatingRowsDefaultCellStyle.BackColor = $uiInput2
            $control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            $control.EnableHeadersVisualStyles = $false
        }
        elseif ($control -is [System.Windows.Forms.Button]) {
            Set-ModernButton -Button $control -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1
        }
        elseif ($control -is [System.Windows.Forms.ProgressBar]) {
            $control.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            $control.BackColor = $uiInput
            $control.ForeColor = $uiAccent2
        }

        if ($control.Controls.Count -gt 0) {
            Set-DarkControlTree -Root $control
        }
    }
}

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(20, 10)
$titleLabel.Size = New-Object System.Drawing.Size(400, 30)
$titleLabel.Text = "Virtual Machine Deletion Tool"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
$titleLabel.ForeColor = $uiText
$form.Controls.Add($titleLabel)

$btnUnlockSettings = New-Object System.Windows.Forms.Button
$btnUnlockSettings.Location = New-Object System.Drawing.Point(552, 8)
$btnUnlockSettings.Size = New-Object System.Drawing.Size(40, 36)
$btnUnlockSettings.Text = "🔒"
$btnUnlockSettings.Font = New-Object System.Drawing.Font("Segoe UI Emoji", 11, [System.Drawing.FontStyle]::Regular)
$btnUnlockSettings.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnUnlockSettings.FlatAppearance.BorderSize = 0
$btnUnlockSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($btnUnlockSettings)

$versionLabel = New-Object System.Windows.Forms.Label
$versionLabel.Location = New-Object System.Drawing.Point(595, 18)
$versionLabel.Size = New-Object System.Drawing.Size(60, 20)
$versionLabel.Text = "v1.0"
$versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$versionLabel.ForeColor = $uiMuted
$versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($versionLabel)
#endregion

#region Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(15, 45)
$tabControl.Size = New-Object System.Drawing.Size(640, 825)
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabControl.ItemSize = New-Object System.Drawing.Size(159, 32)
$tabControl.Add_DrawItem({
    param($sender, $eventArgs)

    $tabPage = $sender.TabPages[$eventArgs.Index]
    $isSelected = ($eventArgs.Index -eq $sender.SelectedIndex)
    $bounds = $sender.GetTabRect($eventArgs.Index)
    $tabRect = New-Object System.Drawing.Rectangle(($bounds.X + 2), ($bounds.Y + 3), ($bounds.Width - 4), ($bounds.Height - 4))

    if ($isSelected) {
        $backColor = $uiAccent2
        $textColor = $uiButtonText
    }
    else {
        $backColor = $uiPanel2
        $textColor = $uiMuted
    }

    $eventArgs.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $tabBackRect = New-Object System.Drawing.Rectangle($bounds.X, $bounds.Y, $bounds.Width, $bounds.Height)
    $tabBackBrush = New-Object System.Drawing.SolidBrush($uiPanel)
    $backBrush = New-Object System.Drawing.SolidBrush($backColor)
    $borderPen = New-Object System.Drawing.Pen($(if ($isSelected) { $uiAccent } else { $uiBorderSoft }), 1)
    $textFlags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor
                 [System.Windows.Forms.TextFormatFlags]::EndEllipsis

    $eventArgs.Graphics.FillRectangle($tabBackBrush, $tabBackRect)
    $eventArgs.Graphics.FillRectangle($backBrush, $tabRect)
    $eventArgs.Graphics.DrawRectangle($borderPen, $tabRect)
    [System.Windows.Forms.TextRenderer]::DrawText($eventArgs.Graphics, $tabPage.Text, $sender.Font, $tabRect, $textColor, $textFlags)

    $tabBackBrush.Dispose()
    $backBrush.Dispose()
    $borderPen.Dispose()
})
$form.Controls.Add($tabControl)

$tabMain = New-Object System.Windows.Forms.TabPage
$tabMain.Text = "VM Deletion"
$tabMain.BackColor = $uiPanel
$tabMain.AutoScroll = $true
$tabControl.Controls.Add($tabMain)

$tabArchive = New-Object System.Windows.Forms.TabPage
$tabArchive.Enabled = $false
$tabArchive.Text = "Archive"
$tabArchive.BackColor = $uiPanel
$tabArchive.AutoScroll = $true
$tabControl.Controls.Add($tabArchive)

$tabSettings = New-Object System.Windows.Forms.TabPage
$tabSettings.Enabled = $false
$tabSettings.Text = "Settings"
$tabSettings.BackColor = $uiPanel
$tabSettings.AutoScroll = $true
$tabControl.Controls.Add($tabSettings)

$tabLogs = New-Object System.Windows.Forms.TabPage
$tabLogs.Text = "LOGS"
$tabLogs.BackColor = $uiPanel
$tabLogs.AutoScroll = $false
$tabControl.Controls.Add($tabLogs)
$script:logsTab = $tabLogs
#endregion

#region Logs Tab Controls
$logsGroup = New-Object System.Windows.Forms.GroupBox
$logsGroup.Location = New-Object System.Drawing.Point(15, 15)
$logsGroup.Size = New-Object System.Drawing.Size(600, 740)
$logsGroup.Text = "Error Details"
$tabLogs.Controls.Add($logsGroup)

$txtLogs = New-Object System.Windows.Forms.TextBox
$txtLogs.Location = New-Object System.Drawing.Point(15, 28)
$txtLogs.Size = New-Object System.Drawing.Size(570, 655)
$txtLogs.Multiline = $true
$txtLogs.ReadOnly = $true
$txtLogs.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$txtLogs.WordWrap = $false
$txtLogs.Font = New-Object System.Drawing.Font("Consolas", 9)
$txtLogs.BackColor = $uiInput
$txtLogs.ForeColor = $uiText
$txtLogs.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$logsGroup.Controls.Add($txtLogs)
$script:txtLogs = $txtLogs

$btnClearLogs = New-Object System.Windows.Forms.Button
$btnClearLogs.Location = New-Object System.Drawing.Point(485, 695)
$btnClearLogs.Size = New-Object System.Drawing.Size(100, 30)
$btnClearLogs.Text = "Clear"
$btnClearLogs.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnClearLogs.FlatStyle = "Flat"
$btnClearLogs.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClearLogs.Add_Click({ $script:txtLogs.Clear() })
$logsGroup.Controls.Add($btnClearLogs)
#endregion
#region Archive Tab Controls

# SQL Connection Settings
$sqlGroup = New-Object System.Windows.Forms.GroupBox
$sqlGroup.Location = New-Object System.Drawing.Point(15, 15)
$sqlGroup.Size = New-Object System.Drawing.Size(600, 130)
$sqlGroup.Text = "SQL Database Connection"
$tabArchive.Controls.Add($sqlGroup)

$lblSQLServer = New-Object System.Windows.Forms.Label
$lblSQLServer.Location = New-Object System.Drawing.Point(15, 28)
$lblSQLServer.Size = New-Object System.Drawing.Size(90, 20)
$lblSQLServer.Text = "SQL Server:"
$sqlGroup.Controls.Add($lblSQLServer)

$txtSQLServer = New-Object System.Windows.Forms.TextBox
$txtSQLServer.Location = New-Object System.Drawing.Point(110, 25)
$txtSQLServer.Size = New-Object System.Drawing.Size(200, 23)
$txtSQLServer.Text = $script:DefaultSettings.SQLServer
$sqlGroup.Controls.Add($txtSQLServer)

$lblSQLDatabase = New-Object System.Windows.Forms.Label
$lblSQLDatabase.Location = New-Object System.Drawing.Point(325, 28)
$lblSQLDatabase.Size = New-Object System.Drawing.Size(80, 20)
$lblSQLDatabase.Text = "Database:"
$sqlGroup.Controls.Add($lblSQLDatabase)

$txtSQLDatabase = New-Object System.Windows.Forms.TextBox
$txtSQLDatabase.Location = New-Object System.Drawing.Point(410, 25)
$txtSQLDatabase.Size = New-Object System.Drawing.Size(160, 23)
$txtSQLDatabase.Text = $script:DefaultSettings.SQLDatabase
$sqlGroup.Controls.Add($txtSQLDatabase)

$lblSQLTable = New-Object System.Windows.Forms.Label
$lblSQLTable.Location = New-Object System.Drawing.Point(15, 58)
$lblSQLTable.Size = New-Object System.Drawing.Size(90, 20)
$lblSQLTable.Text = "Table Name:"
$sqlGroup.Controls.Add($lblSQLTable)

$txtSQLTable = New-Object System.Windows.Forms.TextBox
$txtSQLTable.Location = New-Object System.Drawing.Point(110, 55)
$txtSQLTable.Size = New-Object System.Drawing.Size(200, 23)
$txtSQLTable.Text = $($script:DefaultSettings.SQLTable)
$sqlGroup.Controls.Add($txtSQLTable)

$chkSQLIntegrated = New-Object System.Windows.Forms.CheckBox
$chkSQLIntegrated.Location = New-Object System.Drawing.Point(325, 57)
$chkSQLIntegrated.Size = New-Object System.Drawing.Size(225, 20)
$chkSQLIntegrated.Text = "Windows Authentication"
$chkSQLIntegrated.Checked = $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($chkSQLIntegrated)

$lblSQLUser = New-Object System.Windows.Forms.Label
$lblSQLUser.Location = New-Object System.Drawing.Point(15, 88)
$lblSQLUser.Size = New-Object System.Drawing.Size(85, 20)
$lblSQLUser.Text = "Username:"
$lblSQLUser.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($lblSQLUser)

$txtSQLUser = New-Object System.Windows.Forms.TextBox
$txtSQLUser.Location = New-Object System.Drawing.Point(110, 85)
$txtSQLUser.Size = New-Object System.Drawing.Size(145, 23)
$txtSQLUser.Text = $script:DefaultSettings.SQLUsername
$txtSQLUser.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($txtSQLUser)

$lblSQLPass = New-Object System.Windows.Forms.Label
$lblSQLPass.Location = New-Object System.Drawing.Point(270, 88)
$lblSQLPass.Size = New-Object System.Drawing.Size(80, 20)
$lblSQLPass.Text = "Password:"
$lblSQLPass.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($lblSQLPass)

$txtSQLPass = New-Object System.Windows.Forms.TextBox
$txtSQLPass.Location = New-Object System.Drawing.Point(360, 85)
$txtSQLPass.Size = New-Object System.Drawing.Size(95, 23)
$txtSQLPass.UseSystemPasswordChar = $true
$txtSQLPass.Text = $script:DefaultSettings.SQLPassword
$txtSQLPass.Enabled = -not $script:DefaultSettings.SQLIntegratedAuth
$sqlGroup.Controls.Add($txtSQLPass)

$btnCheckTable = New-Object System.Windows.Forms.Button
$btnCheckTable.Location = New-Object System.Drawing.Point(470, 83)
$btnCheckTable.Size = New-Object System.Drawing.Size(110, 30)
$btnCheckTable.Text = "Check Table"
$btnCheckTable.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
$btnCheckTable.ForeColor = [System.Drawing.Color]::White
$btnCheckTable.FlatStyle = "Flat"
$btnCheckTable.Cursor = [System.Windows.Forms.Cursors]::Hand
$sqlGroup.Controls.Add($btnCheckTable)

# Progress Area
$archiveProgressGroup = New-Object System.Windows.Forms.GroupBox
$archiveProgressGroup.Location = New-Object System.Drawing.Point(15, 155)
$archiveProgressGroup.Size = New-Object System.Drawing.Size(600, 82)
$archiveProgressGroup.Text = "Status"
$tabArchive.Controls.Add($archiveProgressGroup)

$lblArchiveStatus = New-Object System.Windows.Forms.TextBox
$lblArchiveStatus.Location = New-Object System.Drawing.Point(15, 25)
$lblArchiveStatus.Size = New-Object System.Drawing.Size(570, 46)
$lblArchiveStatus.Text = "Click 'Check Table' to verify database connection and table status"
$lblArchiveStatus.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblArchiveStatus.Multiline = $true
$lblArchiveStatus.ReadOnly = $true
$lblArchiveStatus.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$lblArchiveStatus.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$archiveProgressGroup.Controls.Add($lblArchiveStatus)

# Data Grid
$archiveDataGroup = New-Object System.Windows.Forms.GroupBox
$archiveDataGroup.Location = New-Object System.Drawing.Point(15, 247)
$archiveDataGroup.Size = New-Object System.Drawing.Size(600, 453)
$archiveDataGroup.Text = "Deletion Archive Data"
$tabArchive.Controls.Add($archiveDataGroup)

$dgvArchive = New-Object System.Windows.Forms.DataGridView
$dgvArchive.Location = New-Object System.Drawing.Point(10, 20)
$dgvArchive.Size = New-Object System.Drawing.Size(580, 388)
$dgvArchive.AllowUserToAddRows = $false
$dgvArchive.AllowUserToDeleteRows = $false
$dgvArchive.ReadOnly = $true

# ========== שינויים חדשים - גלילה אופקית ==========
$dgvArchive.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::None
$dgvArchive.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$dgvArchive.AllowUserToResizeColumns = $true
# =================================================

$dgvArchive.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgvArchive.BackgroundColor = $uiInput
$dgvArchive.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
$dgvArchive.RowHeadersVisible = $false
$archiveDataGroup.Controls.Add($dgvArchive)

$lblRecordCount = New-Object System.Windows.Forms.Label
$lblRecordCount.Location = New-Object System.Drawing.Point(10, 418)
$lblRecordCount.Size = New-Object System.Drawing.Size(230, 20)
$lblRecordCount.Text = "0 records"
$lblRecordCount.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblRecordCount.ForeColor = $uiMuted
$archiveDataGroup.Controls.Add($lblRecordCount)

$btnRefreshArchive = New-Object System.Windows.Forms.Button
$btnRefreshArchive.Location = New-Object System.Drawing.Point(284, 413)
$btnRefreshArchive.Size = New-Object System.Drawing.Size(90, 30)
$btnRefreshArchive.Text = "Refresh"
$btnRefreshArchive.FlatStyle = "Flat"
$btnRefreshArchive.Enabled = $false
$btnRefreshArchive.Cursor = [System.Windows.Forms.Cursors]::Hand
$archiveDataGroup.Controls.Add($btnRefreshArchive)

$btnExportCSV = New-Object System.Windows.Forms.Button
$btnExportCSV.Location = New-Object System.Drawing.Point(384, 413)
$btnExportCSV.Size = New-Object System.Drawing.Size(92, 30)
$btnExportCSV.Text = "Export CSV"
$btnExportCSV.FlatStyle = "Flat"
$btnExportCSV.Enabled = $false
$btnExportCSV.Cursor = [System.Windows.Forms.Cursors]::Hand
$archiveDataGroup.Controls.Add($btnExportCSV)

$btnExportExcel = New-Object System.Windows.Forms.Button
$btnExportExcel.Location = New-Object System.Drawing.Point(486, 413)
$btnExportExcel.Size = New-Object System.Drawing.Size(104, 30)
$btnExportExcel.Text = "Export Excel"
$btnExportExcel.FlatStyle = "Flat"
$btnExportExcel.Enabled = $false
$btnExportExcel.Cursor = [System.Windows.Forms.Cursors]::Hand
$archiveDataGroup.Controls.Add($btnExportExcel)

$btnSaveSQLSettings = New-Object System.Windows.Forms.Button
$btnSaveSQLSettings.Location = New-Object System.Drawing.Point(510, 715)
$btnSaveSQLSettings.Size = New-Object System.Drawing.Size(105, 30)
$btnSaveSQLSettings.Text = "Save"
$btnSaveSQLSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSaveSQLSettings.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnSaveSQLSettings.ForeColor = [System.Drawing.Color]::White
$btnSaveSQLSettings.FlatStyle = "Flat"
$btnSaveSQLSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$tabArchive.Controls.Add($btnSaveSQLSettings)

#endregion

#region Settings Tab Controls
# SCCM Settings
$sccmSettingsGroup = New-Object System.Windows.Forms.GroupBox
$sccmSettingsGroup.Location = New-Object System.Drawing.Point(15, 15)
$sccmSettingsGroup.Size = New-Object System.Drawing.Size(600, 70)
$sccmSettingsGroup.Text = "SCCM Settings"
$tabSettings.Controls.Add($sccmSettingsGroup)

$lblSCCMServer = New-Object System.Windows.Forms.Label
$lblSCCMServer.Location = New-Object System.Drawing.Point(15, 28)
$lblSCCMServer.Size = New-Object System.Drawing.Size(100, 20)
$lblSCCMServer.Text = "SCCM Server:"
$sccmSettingsGroup.Controls.Add($lblSCCMServer)

$txtSettingsSCCM = New-Object System.Windows.Forms.TextBox
$txtSettingsSCCM.Location = New-Object System.Drawing.Point(120, 25)
$txtSettingsSCCM.Size = New-Object System.Drawing.Size(220, 23)
$txtSettingsSCCM.Text = $script:DefaultSettings.SCCMServer
$sccmSettingsGroup.Controls.Add($txtSettingsSCCM)

$lblSiteCode = New-Object System.Windows.Forms.Label
$lblSiteCode.Location = New-Object System.Drawing.Point(350, 28)
$lblSiteCode.Size = New-Object System.Drawing.Size(45, 20)
$lblSiteCode.Text = "Site:"
$sccmSettingsGroup.Controls.Add($lblSiteCode)

$txtSettingsSiteCode = New-Object System.Windows.Forms.TextBox
$txtSettingsSiteCode.Location = New-Object System.Drawing.Point(405, 25)
$txtSettingsSiteCode.Size = New-Object System.Drawing.Size(55, 23)
$txtSettingsSiteCode.MaxLength = 3
$txtSettingsSiteCode.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper
$txtSettingsSiteCode.Text = $script:DefaultSettings.SCCMSiteCode
$sccmSettingsGroup.Controls.Add($txtSettingsSiteCode)

# DNS Settings
$dnsSettingsGroup = New-Object System.Windows.Forms.GroupBox
$dnsSettingsGroup.Location = New-Object System.Drawing.Point(15, 95)
$dnsSettingsGroup.Size = New-Object System.Drawing.Size(600, 70)
$dnsSettingsGroup.Text = "DNS Settings"
$tabSettings.Controls.Add($dnsSettingsGroup)

$lblDNSServer = New-Object System.Windows.Forms.Label
$lblDNSServer.Location = New-Object System.Drawing.Point(15, 28)
$lblDNSServer.Size = New-Object System.Drawing.Size(90, 20)
$lblDNSServer.Text = "DNS Server:"
$dnsSettingsGroup.Controls.Add($lblDNSServer)

$txtSettingsDNS = New-Object System.Windows.Forms.TextBox
$txtSettingsDNS.Location = New-Object System.Drawing.Point(115, 25)
$txtSettingsDNS.Size = New-Object System.Drawing.Size(160, 23)
$txtSettingsDNS.Text = $script:DefaultSettings.DNSServer
$dnsSettingsGroup.Controls.Add($txtSettingsDNS)

$lblDNSZone = New-Object System.Windows.Forms.Label
$lblDNSZone.Location = New-Object System.Drawing.Point(285, 28)
$lblDNSZone.Size = New-Object System.Drawing.Size(45, 20)
$lblDNSZone.Text = "DNS:"
$dnsSettingsGroup.Controls.Add($lblDNSZone)

$txtSettingsDNSZone = New-Object System.Windows.Forms.TextBox
$txtSettingsDNSZone.Location = New-Object System.Drawing.Point(335, 25)
$txtSettingsDNSZone.Size = New-Object System.Drawing.Size(135, 23)
$txtSettingsDNSZone.Text = $script:DefaultSettings.DNSZone
$dnsSettingsGroup.Controls.Add($txtSettingsDNSZone)

# AD Settings
$adSettingsGroup = New-Object System.Windows.Forms.GroupBox
$adSettingsGroup.Location = New-Object System.Drawing.Point(15, 175)
$adSettingsGroup.Size = New-Object System.Drawing.Size(600, 100)
$adSettingsGroup.Text = "Active Directory Settings"
$tabSettings.Controls.Add($adSettingsGroup)

$lblDC = New-Object System.Windows.Forms.Label
$lblDC.Location = New-Object System.Drawing.Point(15, 28)
$lblDC.Size = New-Object System.Drawing.Size(135, 20)
$lblDC.Text = "Domain Controller:"
$adSettingsGroup.Controls.Add($lblDC)

$txtSettingsDC = New-Object System.Windows.Forms.TextBox
$txtSettingsDC.Location = New-Object System.Drawing.Point(165, 25)
$txtSettingsDC.Size = New-Object System.Drawing.Size(300, 23)
$txtSettingsDC.Text = $script:DefaultSettings.DomainController
$adSettingsGroup.Controls.Add($txtSettingsDC)

$lblEmail = New-Object System.Windows.Forms.Label
$lblEmail.Location = New-Object System.Drawing.Point(15, 58)
$lblEmail.Size = New-Object System.Drawing.Size(150, 20)
$lblEmail.Text = "Notification Email:"
$adSettingsGroup.Controls.Add($lblEmail)

$txtSettingsEmail = New-Object System.Windows.Forms.TextBox
$txtSettingsEmail.Location = New-Object System.Drawing.Point(175, 55)
$txtSettingsEmail.Size = New-Object System.Drawing.Size(285, 23)
$txtSettingsEmail.Text = $script:DefaultSettings.EmailNotification
$adSettingsGroup.Controls.Add($txtSettingsEmail)


# General Options
$optionsGroup = New-Object System.Windows.Forms.GroupBox
$optionsGroup.Location = New-Object System.Drawing.Point(15, 285)
$optionsGroup.Size = New-Object System.Drawing.Size(600, 130)
$optionsGroup.Text = "General Options"
$tabSettings.Controls.Add($optionsGroup)

$chkEnableLogging = New-Object System.Windows.Forms.CheckBox
$chkEnableLogging.Location = New-Object System.Drawing.Point(15, 25)
$chkEnableLogging.Size = New-Object System.Drawing.Size(150, 20)
$chkEnableLogging.Text = "Enable Logging"
$chkEnableLogging.Checked = $script:DefaultSettings.EnableLogging
$optionsGroup.Controls.Add($chkEnableLogging)

$chkConfirmDeletion = New-Object System.Windows.Forms.CheckBox
$chkConfirmDeletion.Location = New-Object System.Drawing.Point(200, 25)
$chkConfirmDeletion.Size = New-Object System.Drawing.Size(230, 20)
$chkConfirmDeletion.Text = "Require Deletion Confirmation"
$chkConfirmDeletion.Checked = $script:DefaultSettings.ConfirmDeletion
$optionsGroup.Controls.Add($chkConfirmDeletion)

$chkDeleteFromDisk = New-Object System.Windows.Forms.CheckBox
$chkDeleteFromDisk.Location = New-Object System.Drawing.Point(15, 50)
$chkDeleteFromDisk.Size = New-Object System.Drawing.Size(280, 20)
$chkDeleteFromDisk.Text = "Delete VM files from disk permanently"
$chkDeleteFromDisk.Checked = $script:DefaultSettings.DeleteFromDisk
$optionsGroup.Controls.Add($chkDeleteFromDisk)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Location = New-Object System.Drawing.Point(320, 50)
$lblTimeout.Size = New-Object System.Drawing.Size(135, 20)
$lblTimeout.Text = "Power Off Timeout:"
$optionsGroup.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(465, 48)
$numTimeout.Size = New-Object System.Drawing.Size(60, 23)
$numTimeout.Minimum = 10
$numTimeout.Maximum = 300
$numTimeout.Value = $script:DefaultSettings.PowerOffTimeout
$optionsGroup.Controls.Add($numTimeout)

$lblLogPath = New-Object System.Windows.Forms.Label
$lblLogPath.Location = New-Object System.Drawing.Point(15, 85)
$lblLogPath.Size = New-Object System.Drawing.Size(80, 20)
$lblLogPath.Text = "Log Path:"
$optionsGroup.Controls.Add($lblLogPath)

$txtLogPath = New-Object System.Windows.Forms.TextBox
$txtLogPath.Location = New-Object System.Drawing.Point(105, 82)
$txtLogPath.Size = New-Object System.Drawing.Size(370, 23)
$txtLogPath.Text = $script:DefaultSettings.LogPath
$optionsGroup.Controls.Add($txtLogPath)

$btnBrowseLog = New-Object System.Windows.Forms.Button
$btnBrowseLog.Location = New-Object System.Drawing.Point(485, 79)
$btnBrowseLog.Size = New-Object System.Drawing.Size(85, 30)
$btnBrowseLog.Text = "Browse..."
$btnBrowseLog.FlatStyle = "Flat"
$optionsGroup.Controls.Add($btnBrowseLog)

# Email Settings Group
$emailSettingsGroup = New-Object System.Windows.Forms.GroupBox
$emailSettingsGroup.Location = New-Object System.Drawing.Point(15, 425)
$emailSettingsGroup.Size = New-Object System.Drawing.Size(600, 275)
$emailSettingsGroup.Text = "Email Notification Settings"
$tabSettings.Controls.Add($emailSettingsGroup)

# Enable Email
$chkEnableEmail = New-Object System.Windows.Forms.CheckBox
$chkEnableEmail.Location = New-Object System.Drawing.Point(15, 22)
$chkEnableEmail.Size = New-Object System.Drawing.Size(205, 20)
$chkEnableEmail.Text = "Enable Email Notifications"
$chkEnableEmail.Checked = $script:DefaultSettings.EnableEmailNotification
$emailSettingsGroup.Controls.Add($chkEnableEmail)

# SMTP Row
$lblSMTPServer = New-Object System.Windows.Forms.Label
$lblSMTPServer.Location = New-Object System.Drawing.Point(15, 50)
$lblSMTPServer.Size = New-Object System.Drawing.Size(45, 20)
$lblSMTPServer.Text = "SMTP:"
$emailSettingsGroup.Controls.Add($lblSMTPServer)

$txtSMTPServer = New-Object System.Windows.Forms.TextBox
$txtSMTPServer.Location = New-Object System.Drawing.Point(65, 47)
$txtSMTPServer.Size = New-Object System.Drawing.Size(130, 23)
$txtSMTPServer.Text = $script:DefaultSettings.SMTPServer
$emailSettingsGroup.Controls.Add($txtSMTPServer)

$lblSMTPPort = New-Object System.Windows.Forms.Label
$lblSMTPPort.Location = New-Object System.Drawing.Point(200, 50)
$lblSMTPPort.Size = New-Object System.Drawing.Size(45, 20)
$lblSMTPPort.Text = "Port:"
$emailSettingsGroup.Controls.Add($lblSMTPPort)

$numSMTPPort = New-Object System.Windows.Forms.NumericUpDown
$numSMTPPort.Location = New-Object System.Drawing.Point(250, 47)
$numSMTPPort.Size = New-Object System.Drawing.Size(55, 23)
$numSMTPPort.Minimum = 1
$numSMTPPort.Maximum = 65535
$numSMTPPort.Value = $script:DefaultSettings.SMTPPort
$emailSettingsGroup.Controls.Add($numSMTPPort)

$chkSMTPSSL = New-Object System.Windows.Forms.CheckBox
$chkSMTPSSL.Location = New-Object System.Drawing.Point(315, 48)
$chkSMTPSSL.Size = New-Object System.Drawing.Size(55, 20)
$chkSMTPSSL.Text = "SSL"
$chkSMTPSSL.Checked = $script:DefaultSettings.SMTPUseSSL
$emailSettingsGroup.Controls.Add($chkSMTPSSL)

$lblSMTPFrom = New-Object System.Windows.Forms.Label
$lblSMTPFrom.Location = New-Object System.Drawing.Point(380, 50)
$lblSMTPFrom.Size = New-Object System.Drawing.Size(45, 20)
$lblSMTPFrom.Text = "From:"
$emailSettingsGroup.Controls.Add($lblSMTPFrom)

$txtSMTPFrom = New-Object System.Windows.Forms.TextBox
$txtSMTPFrom.Location = New-Object System.Drawing.Point(430, 47)
$txtSMTPFrom.Size = New-Object System.Drawing.Size(140, 23)
$txtSMTPFrom.Text = $script:DefaultSettings.SMTPFrom
$emailSettingsGroup.Controls.Add($txtSMTPFrom)

# Groups Section - Headers
$lblGroupsCommon = New-Object System.Windows.Forms.Label
$lblGroupsCommon.Location = New-Object System.Drawing.Point(15, 80)
$lblGroupsCommon.Size = New-Object System.Drawing.Size(180, 15)
$lblGroupsCommon.Text = "Other Groups (All VMs):"
$lblGroupsCommon.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$emailSettingsGroup.Controls.Add($lblGroupsCommon)

$lblGroupsWindows = New-Object System.Windows.Forms.Label
$lblGroupsWindows.Location = New-Object System.Drawing.Point(210, 80)
$lblGroupsWindows.Size = New-Object System.Drawing.Size(180, 15)
$lblGroupsWindows.Text = "Windows Only:"
$lblGroupsWindows.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblGroupsWindows.ForeColor = $uiAccent
$emailSettingsGroup.Controls.Add($lblGroupsWindows)

$lblGroupsLinux = New-Object System.Windows.Forms.Label
$lblGroupsLinux.Location = New-Object System.Drawing.Point(405, 80)
$lblGroupsLinux.Size = New-Object System.Drawing.Size(180, 15)
$lblGroupsLinux.Text = "Linux Only:"
$lblGroupsLinux.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblGroupsLinux.ForeColor = $uiWarning
$emailSettingsGroup.Controls.Add($lblGroupsLinux)

# ========== ListBoxes for Groups ==========

# Other Groups ListBox
$script:lstGroupsCommon = New-Object System.Windows.Forms.ListBox
$script:lstGroupsCommon.Name = "lstGroupsCommon"
$script:lstGroupsCommon.Location = New-Object System.Drawing.Point(15, 98)
$script:lstGroupsCommon.Size = New-Object System.Drawing.Size(180, 60)
$script:lstGroupsCommon.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lstGroupsCommon.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$script:lstGroupsCommon.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$emailSettingsGroup.Controls.Add($script:lstGroupsCommon)

# Windows Groups ListBox
$script:lstGroupsWindows = New-Object System.Windows.Forms.ListBox
$script:lstGroupsWindows.Name = "lstGroupsWindows"
$script:lstGroupsWindows.Location = New-Object System.Drawing.Point(210, 98)
$script:lstGroupsWindows.Size = New-Object System.Drawing.Size(180, 60)
$script:lstGroupsWindows.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lstGroupsWindows.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$script:lstGroupsWindows.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$emailSettingsGroup.Controls.Add($script:lstGroupsWindows)

# Linux Groups ListBox
$script:lstGroupsLinux = New-Object System.Windows.Forms.ListBox
$script:lstGroupsLinux.Name = "lstGroupsLinux"
$script:lstGroupsLinux.Location = New-Object System.Drawing.Point(405, 98)
$script:lstGroupsLinux.Size = New-Object System.Drawing.Size(180, 60)
$script:lstGroupsLinux.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$script:lstGroupsLinux.SelectionMode = [System.Windows.Forms.SelectionMode]::One
$script:lstGroupsLinux.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$emailSettingsGroup.Controls.Add($script:lstGroupsLinux)

# Add Group Controls
$lblAddGroup = New-Object System.Windows.Forms.Label
$lblAddGroup.Location = New-Object System.Drawing.Point(15, 202)
$lblAddGroup.Size = New-Object System.Drawing.Size(75, 20)
$lblAddGroup.Text = "Add:"
$emailSettingsGroup.Controls.Add($lblAddGroup)

$txtNewGroup = New-Object System.Windows.Forms.TextBox
$txtNewGroup.Location = New-Object System.Drawing.Point(90, 199)
$txtNewGroup.Size = New-Object System.Drawing.Size(245, 23)
$emailSettingsGroup.Controls.Add($txtNewGroup)

$cboGroupType = New-Object System.Windows.Forms.ComboBox
$cboGroupType.Location = New-Object System.Drawing.Point(345, 199)
$cboGroupType.Size = New-Object System.Drawing.Size(110, 23)
$cboGroupType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboGroupType.Items.AddRange(@("Other", "Windows", "Linux"))
$cboGroupType.SelectedIndex = 0
$emailSettingsGroup.Controls.Add($cboGroupType)

$btnAddGroup = New-Object System.Windows.Forms.Button
$btnAddGroup.Location = New-Object System.Drawing.Point(285, 230)
$btnAddGroup.Size = New-Object System.Drawing.Size(80, 30)
$btnAddGroup.Text = "Add"
$btnAddGroup.FlatStyle = "Flat"
$btnAddGroup.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnAddGroup.ForeColor = [System.Drawing.Color]::White
$btnAddGroup.Cursor = [System.Windows.Forms.Cursors]::Hand
$emailSettingsGroup.Controls.Add($btnAddGroup)

$btnRemoveGroup = New-Object System.Windows.Forms.Button
$btnRemoveGroup.Location = New-Object System.Drawing.Point(375, 230)
$btnRemoveGroup.Size = New-Object System.Drawing.Size(95, 30)
$btnRemoveGroup.Text = "Remove"
$btnRemoveGroup.FlatStyle = "Flat"
$btnRemoveGroup.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$btnRemoveGroup.ForeColor = [System.Drawing.Color]::White
$btnRemoveGroup.Cursor = [System.Windows.Forms.Cursors]::Hand
$emailSettingsGroup.Controls.Add($btnRemoveGroup)

$btnTestEmail = New-Object System.Windows.Forms.Button
$btnTestEmail.Location = New-Object System.Drawing.Point(480, 230)
$btnTestEmail.Size = New-Object System.Drawing.Size(105, 30)
$btnTestEmail.Text = "Test Email"
$btnTestEmail.FlatStyle = "Flat"
$btnTestEmail.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
$btnTestEmail.ForeColor = [System.Drawing.Color]::White
$btnTestEmail.Cursor = [System.Windows.Forms.Cursors]::Hand
$emailSettingsGroup.Controls.Add($btnTestEmail)

# ========== טעינת קבוצות קיימות ==========
# Common Groups
if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsCommon)) {
    $groups = $script:DefaultSettings.GroupsCommon.Split(',')
    foreach ($group in $groups) {
        $g = $group.Trim()
        if (-not [string]::IsNullOrWhiteSpace($g)) {
            [void]$script:lstGroupsCommon.Items.Add($g)
        }
    }
}

# Windows Groups
if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsWindows)) {
    $groups = $script:DefaultSettings.GroupsWindows.Split(',')
    foreach ($group in $groups) {
        $g = $group.Trim()
        if (-not [string]::IsNullOrWhiteSpace($g)) {
            [void]$script:lstGroupsWindows.Items.Add($g)
        }
    }
}

# Linux Groups
if (-not [string]::IsNullOrWhiteSpace($script:DefaultSettings.GroupsLinux)) {
    $groups = $script:DefaultSettings.GroupsLinux.Split(',')
    foreach ($group in $groups) {
        $g = $group.Trim()
        if (-not [string]::IsNullOrWhiteSpace($g)) {
            [void]$script:lstGroupsLinux.Items.Add($g)
        }
    }
}


# Settings Buttons
$btnSaveSettings = New-Object System.Windows.Forms.Button
$btnSaveSettings.Location = New-Object System.Drawing.Point(510, 715)
$btnSaveSettings.Size = New-Object System.Drawing.Size(100, 30)
$btnSaveSettings.Text = "Save"
$btnSaveSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSaveSettings.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnSaveSettings.ForeColor = [System.Drawing.Color]::White
$btnSaveSettings.FlatStyle = "Flat"
$btnSaveSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$tabSettings.Controls.Add($btnSaveSettings)

$btnResetSettings = New-Object System.Windows.Forms.Button
$btnResetSettings.Location = New-Object System.Drawing.Point(380, 715)
$btnResetSettings.Size = New-Object System.Drawing.Size(120, 30)
$btnResetSettings.Text = "Reset Settings"
$btnResetSettings.FlatStyle = "Flat"
$btnResetSettings.BackColor = [System.Drawing.Color]::FromArgb(255, 193, 7)
$btnResetSettings.Cursor = [System.Windows.Forms.Cursors]::Hand
$tabSettings.Controls.Add($btnResetSettings)

$lblSettingsStatus = New-Object System.Windows.Forms.Label
$lblSettingsStatus.Location = New-Object System.Drawing.Point(15, 720)
$lblSettingsStatus.Size = New-Object System.Drawing.Size(350, 20)
$lblSettingsStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$tabSettings.Controls.Add($lblSettingsStatus)
#endregion

#region Main Tab Controls

# vCenter Connection
$vcGroup = New-Object System.Windows.Forms.GroupBox
$vcGroup.Location = New-Object System.Drawing.Point(15, 15)
$vcGroup.Size = New-Object System.Drawing.Size(600, 105)
$vcGroup.Text = "vCenter Connection"
$tabMain.Controls.Add($vcGroup)

$lblVC = New-Object System.Windows.Forms.Label
$lblVC.Location = New-Object System.Drawing.Point(15, 28)
$lblVC.Size = New-Object System.Drawing.Size(115, 20)
$lblVC.Text = "vCenter Server:"
$vcGroup.Controls.Add($lblVC)

$cboVCServer = New-Object System.Windows.Forms.ComboBox
$cboVCServer.Location = New-Object System.Drawing.Point(135, 25)
$cboVCServer.Size = New-Object System.Drawing.Size(325, 25)
$cboVCServer.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboVCServer.Font = New-Object System.Drawing.Font("Segoe UI", 9)

foreach ($vc in $script:vCenterServers) {
    $cboVCServer.Items.Add($vc.Name.ToUpper()) | Out-Null
}

$selectedIndex = 0
    for ($i = 0; $i -lt $script:vCenterServers.Count; $i++) {
    if ($script:vCenterServers[$i].Server -eq $script:Settings.vCenterServer) {
        $selectedIndex = $i
        break
    }
}
$cboVCServer.SelectedIndex = $selectedIndex
$vcGroup.Controls.Add($cboVCServer)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(475, 22)
$btnConnect.Size = New-Object System.Drawing.Size(110, 30)
$btnConnect.Text = "Connect"
Set-ModernButton -Button $btnConnect -BackColor $uiAccent -ForeColor ([System.Drawing.Color]::White)
$vcGroup.Controls.Add($btnConnect)

$lblConnectionStatus = New-Object System.Windows.Forms.Label
$lblConnectionStatus.Location = New-Object System.Drawing.Point(15, 55)
$lblConnectionStatus.Size = New-Object System.Drawing.Size(570, 36)
$lblConnectionStatus.Text = "Not connected"
$lblConnectionStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblConnectionStatus.ForeColor = $uiMuted
$vcGroup.Controls.Add($lblConnectionStatus)
Set-ModernGroup -Group $vcGroup

# VM Selection
$vmGroup = New-Object System.Windows.Forms.GroupBox
$vmGroup.Location = New-Object System.Drawing.Point(15, 130)
$vmGroup.Size = New-Object System.Drawing.Size(600, 235)
$vmGroup.Text = "Virtual Machine Selection (Hold Ctrl/Shift for multi-select)"
$tabMain.Controls.Add($vmGroup)

# === שורת הפילטרים ===
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Location = New-Object System.Drawing.Point(15, 25)
$lblFilter.Size = New-Object System.Drawing.Size(60, 20)
$lblFilter.Text = "Filter:"
$vmGroup.Controls.Add($lblFilter)

$txtFilter = New-Object System.Windows.Forms.TextBox
$txtFilter.Location = New-Object System.Drawing.Point(80, 22)
$txtFilter.Size = New-Object System.Drawing.Size(95, 23)
$txtFilter.Enabled = $false
$vmGroup.Controls.Add($txtFilter)

$lblOSFilter = New-Object System.Windows.Forms.Label
$lblOSFilter.Location = New-Object System.Drawing.Point(185, 25)
$lblOSFilter.Size = New-Object System.Drawing.Size(35, 20)
$lblOSFilter.Text = "OS:"
$vmGroup.Controls.Add($lblOSFilter)

$cboOSFilter = New-Object System.Windows.Forms.ComboBox
$cboOSFilter.Location = New-Object System.Drawing.Point(225, 22)
$cboOSFilter.Size = New-Object System.Drawing.Size(90, 23)
$cboOSFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboOSFilter.Items.AddRange(@("All", "Windows", "Linux", "Other"))
$cboOSFilter.SelectedIndex = 0
$cboOSFilter.Enabled = $false
$vmGroup.Controls.Add($cboOSFilter)

$lblPowerFilter = New-Object System.Windows.Forms.Label
$lblPowerFilter.Location = New-Object System.Drawing.Point(320, 25)
$lblPowerFilter.Size = New-Object System.Drawing.Size(45, 20)
$lblPowerFilter.Text = "Power:"
$vmGroup.Controls.Add($lblPowerFilter)

$cboPowerFilter = New-Object System.Windows.Forms.ComboBox
$cboPowerFilter.Location = New-Object System.Drawing.Point(370, 22)
$cboPowerFilter.Size = New-Object System.Drawing.Size(90, 23)
$cboPowerFilter.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cboPowerFilter.Items.AddRange(@("All", "On", "Off", "Suspended"))
$cboPowerFilter.SelectedIndex = 0
$cboPowerFilter.Enabled = $false
$vmGroup.Controls.Add($cboPowerFilter)

# Include Templates Checkbox
$chkIncludeTemplates = New-Object System.Windows.Forms.CheckBox
$chkIncludeTemplates.Location = New-Object System.Drawing.Point(505, 177)
$chkIncludeTemplates.Size = New-Object System.Drawing.Size(85, 20)
$chkIncludeTemplates.Text = "Templates"
$chkIncludeTemplates.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$chkIncludeTemplates.ForeColor = $uiAccent
$chkIncludeTemplates.Checked = $false
$chkIncludeTemplates.Enabled = $false
$vmGroup.Controls.Add($chkIncludeTemplates)

$lblVMCount = New-Object System.Windows.Forms.Label
$lblVMCount.Location = New-Object System.Drawing.Point(465, 25)
$lblVMCount.Size = New-Object System.Drawing.Size(65, 20)
$lblVMCount.Text = "0 VMs"
$lblVMCount.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$lblVMCount.ForeColor = $uiMuted
$vmGroup.Controls.Add($lblVMCount)

$lblSelectedCount = New-Object System.Windows.Forms.Label
$lblSelectedCount.Location = New-Object System.Drawing.Point(530, 25)
$lblSelectedCount.Size = New-Object System.Drawing.Size(65, 20)
$lblSelectedCount.Text = "0 VM selected"
$lblSelectedCount.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$lblSelectedCount.ForeColor = $uiAccent
$vmGroup.Controls.Add($lblSelectedCount)

# === ListView - עם CheckBoxes ===
$lstVMs = New-Object System.Windows.Forms.ListView
$lstVMs.Location = New-Object System.Drawing.Point(15, 50)
$lstVMs.Size = New-Object System.Drawing.Size(495, 120)
$lstVMs.View = [System.Windows.Forms.View]::Details
$lstVMs.FullRowSelect = $true
$lstVMs.GridLines = $true
$lstVMs.CheckBoxes = $true  # הפעלת CheckBoxes
$lstVMs.MultiSelect = $true
$lstVMs.HideSelection = $false
$lstVMs.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lstVMs.Enabled = $true
$lstVMs.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$lstVMs.BackColor = $uiInput
$lstVMs.ForeColor = $uiText

# עמודות
$colName = New-Object System.Windows.Forms.ColumnHeader
$colName.Text = "VM Name"
$colName.Width = 180

$colPowerState = New-Object System.Windows.Forms.ColumnHeader
$colPowerState.Text = "Power State"
$colPowerState.Width = 80

$colOS = New-Object System.Windows.Forms.ColumnHeader
$colOS.Text = "OS Type"
$colOS.Width = 70

$colOSFull = New-Object System.Windows.Forms.ColumnHeader
$colOSFull.Text = "Operating System"
$colOSFull.Width = 165

$lstVMs.Columns.AddRange(@($colName, $colPowerState, $colOS, $colOSFull))
$vmGroup.Controls.Add($lstVMs)

# === כפתורים ליד הטבלה ===

# Refresh Button
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Location = New-Object System.Drawing.Point(520, 50)
$btnRefresh.Size = New-Object System.Drawing.Size(70, 30)
$btnRefresh.Text = "Refresh"
$btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnRefresh.ForeColor = $uiSuccess
$btnRefresh.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnRefresh.FlatAppearance.BorderSize = 1
$btnRefresh.FlatAppearance.BorderColor = $uiSuccess
$btnRefresh.Enabled = $false
$btnRefresh.Cursor = [System.Windows.Forms.Cursors]::Hand
$vmGroup.Controls.Add($btnRefresh)

$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($btnRefresh, "Refresh VM List")

# Clear Selection Button
$btnClearSelection = New-Object System.Windows.Forms.Button
$btnClearSelection.Location = New-Object System.Drawing.Point(520, 88)
$btnClearSelection.Size = New-Object System.Drawing.Size(70, 30)
$btnClearSelection.Text = "Clear"
$btnClearSelection.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnClearSelection.ForeColor = $uiDanger
$btnClearSelection.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnClearSelection.FlatAppearance.BorderSize = 1
$btnClearSelection.FlatAppearance.BorderColor = $uiDanger
$btnClearSelection.Enabled = $false
$btnClearSelection.Cursor = [System.Windows.Forms.Cursors]::Hand
$vmGroup.Controls.Add($btnClearSelection)

$tooltip.SetToolTip($btnClearSelection, "Clear Selection")

# Load from CSV Button
$btnLoadCSV = New-Object System.Windows.Forms.Button
$btnLoadCSV.Location = New-Object System.Drawing.Point(520, 126)
$btnLoadCSV.Size = New-Object System.Drawing.Size(70, 30)
$btnLoadCSV.Text = "CSV"
$btnLoadCSV.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnLoadCSV.ForeColor = $uiAccent
$btnLoadCSV.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLoadCSV.FlatAppearance.BorderSize = 1
$btnLoadCSV.FlatAppearance.BorderColor = $uiAccent
$btnLoadCSV.Enabled = $false
$btnLoadCSV.Cursor = [System.Windows.Forms.Cursors]::Hand
$vmGroup.Controls.Add($btnLoadCSV)

$tooltip.SetToolTip($btnLoadCSV, "Load VM list from CSV file")

$tooltip.SetToolTip($chkIncludeTemplates, "Include VM Templates in the list")


# === Reason - שורה תחתונה ===
$lblReason = New-Object System.Windows.Forms.Label
$lblReason.Location = New-Object System.Drawing.Point(15, 180)
$lblReason.Size = New-Object System.Drawing.Size(65, 20)
$lblReason.Text = "Reason:"
$vmGroup.Controls.Add($lblReason)

$txtReason = New-Object System.Windows.Forms.TextBox
$txtReason.Location = New-Object System.Drawing.Point(90, 177)
$txtReason.Size = New-Object System.Drawing.Size(400, 23)
$txtReason.Enabled = $false
$txtReason.MaxLength = 500
$txtReason.Text = "Activity Ended"
$vmGroup.Controls.Add($txtReason)

$lblReasonHint = New-Object System.Windows.Forms.Label
$lblReasonHint.Location = New-Object System.Drawing.Point(90, 202)
$lblReasonHint.Size = New-Object System.Drawing.Size(380, 18)
$lblReasonHint.Text = "Enter reason for deletion (saved to archive &email)"
$lblReasonHint.Font = New-Object System.Drawing.Font("Segoe UI", 7)
$lblReasonHint.ForeColor = $uiMuted
$vmGroup.Controls.Add($lblReasonHint)
Set-ModernGroup -Group $vmGroup

# Deletion Options
$deleteOptionsGroup = New-Object System.Windows.Forms.GroupBox
$deleteOptionsGroup.Location = New-Object System.Drawing.Point(15, 375)
$deleteOptionsGroup.Size = New-Object System.Drawing.Size(600, 55)
$deleteOptionsGroup.Text = "Deletion Options"
$tabMain.Controls.Add($deleteOptionsGroup)

$chkSCCM = New-Object System.Windows.Forms.CheckBox
$chkSCCM.Location = New-Object System.Drawing.Point(20, 20)
$chkSCCM.Size = New-Object System.Drawing.Size(80, 22)
$chkSCCM.Text = "SCCM"
$chkSCCM.Checked = $true
$deleteOptionsGroup.Controls.Add($chkSCCM)

$chkAD = New-Object System.Windows.Forms.CheckBox
$chkAD.Location = New-Object System.Drawing.Point(115, 20)
$chkAD.Size = New-Object System.Drawing.Size(80, 22)
$chkAD.Text = "AD"
$chkAD.Checked = $true
$deleteOptionsGroup.Controls.Add($chkAD)

$chkDNS = New-Object System.Windows.Forms.CheckBox
$chkDNS.Location = New-Object System.Drawing.Point(210, 20)
$chkDNS.Size = New-Object System.Drawing.Size(80, 22)
$chkDNS.Text = "DNS"
$chkDNS.Checked = $true
$deleteOptionsGroup.Controls.Add($chkDNS)

$chkVM = New-Object System.Windows.Forms.CheckBox
$chkVM.Location = New-Object System.Drawing.Point(305, 20)
$chkVM.Size = New-Object System.Drawing.Size(80, 22)
$chkVM.Text = "VM"
$chkVM.Checked = $true
$deleteOptionsGroup.Controls.Add($chkVM)

$chkArchive = New-Object System.Windows.Forms.CheckBox
$chkArchive.Location = New-Object System.Drawing.Point(400, 20)
$chkArchive.Size = New-Object System.Drawing.Size(80, 22)
$chkArchive.Text = "Archive"
$chkArchive.Checked = $true
$deleteOptionsGroup.Controls.Add($chkArchive)

$chkEmail = New-Object System.Windows.Forms.CheckBox
$chkEmail.Location = New-Object System.Drawing.Point(495, 20)
$chkEmail.Size = New-Object System.Drawing.Size(80, 22)
$chkEmail.Text = "Email"
$chkEmail.Checked = $true
$deleteOptionsGroup.Controls.Add($chkEmail)
Set-ModernGroup -Group $deleteOptionsGroup

# Progress
$progressGroup = New-Object System.Windows.Forms.GroupBox
$progressGroup.Location = New-Object System.Drawing.Point(15, 440)
$progressGroup.Size = New-Object System.Drawing.Size(600, 260)
$progressGroup.Text = "Deletion Progress"
$tabMain.Controls.Add($progressGroup)


$lblVMProgress = New-Object System.Windows.Forms.Label
$lblVMProgress.Location = New-Object System.Drawing.Point(10, 25)
$lblVMProgress.Size = New-Object System.Drawing.Size(500, 20)
$lblVMProgress.Text = "VM Progress: Idle"
$lblVMProgress.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblVMProgress.ForeColor = $uiAccent

$progressGroup.Controls.Add($lblVMProgress)


foreach ($task in $tasks) {
    $idx = $tasks.IndexOf($task)
    
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(15, $yPos)
    $statusLabel.Size = New-Object System.Drawing.Size(46, 22)
    $statusLabel.Text = "WAIT"
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $statusLabel.ForeColor = $uiMuted
    $progressGroup.Controls.Add($statusLabel)
    $statusLabels += $statusLabel
    
    $taskLabel = New-Object System.Windows.Forms.Label
    $taskLabel.Location = New-Object System.Drawing.Point(68, $yPos)
    $taskLabel.Size = New-Object System.Drawing.Size(177, 22)
    $taskLabel.Text = $task
    $taskLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $taskLabel.ForeColor = $uiText
    $progressGroup.Controls.Add($taskLabel)
    
    $resultLabel = New-Object System.Windows.Forms.Label
    $resultLabel.Location = New-Object System.Drawing.Point(255, $yPos)
    $resultLabel.Size = New-Object System.Drawing.Size(330, 22)
    $resultLabel.Text = "Pending"
    $resultLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $resultLabel.ForeColor = $uiMuted
    $progressGroup.Controls.Add($resultLabel)
    $resultLabels += $resultLabel
    
    $yPos += 25
}

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(15, 225)
$progressBar.Size = New-Object System.Drawing.Size(570, 20)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressGroup.Controls.Add($progressBar)
Set-ModernGroup -Group $progressGroup

# Action Buttons
$btnDelete = New-Object System.Windows.Forms.Button
$btnDelete.Location = New-Object System.Drawing.Point(510, 715)
$btnDelete.Size = New-Object System.Drawing.Size(100, 30)
$btnDelete.Text = "Delete VM"
$btnDelete.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnDelete.Enabled = $false
Set-ModernButton -Button $btnDelete -BackColor $uiDanger -ForeColor ([System.Drawing.Color]::White)
$tabMain.Controls.Add($btnDelete)

# כפתור Resume
$btnResume = New-Object System.Windows.Forms.Button
$btnResume.Location = New-Object System.Drawing.Point(400, 715)
$btnResume.Size = New-Object System.Drawing.Size(100, 30)
$btnResume.Text = "Resume"
$btnResume.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
Set-ModernButton -Button $btnResume -BackColor $uiWarning -ForeColor $uiText
$tabMain.Controls.Add($btnResume)

# בדוק אם יש קובץ progress והפעל את הכפתור בהתאם
$progressFile = "$env:TEMP\VMDeletionTool\progress.json"
$btnResume.Enabled = [bool](Wait-JsonFileAvailable -Path $progressFile)

$tooltip.SetToolTip($btnResume, "Resume from last failed VM")

# כפתור Close
$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-Object System.Drawing.Point(510, 675)
$btnClose.Size = New-Object System.Drawing.Size(100, 30)
$btnClose.Text = "Close"
$btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnClose.FlatStyle = "Flat"
$btnClose.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnClose.Visible = $false

Set-DarkControlTree -Root $form
Set-ButtonPaddingTree -Root $form
Set-AccessibilityTree -Root $form
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.BackColor = [System.Drawing.Color]::Transparent
$versionLabel.ForeColor = $uiMuted
$lblVMCount.ForeColor = $uiMuted
$lblSelectedCount.ForeColor = $uiAccent
$lblReasonHint.ForeColor = $uiMuted
$lstVMs.BackColor = $uiInput
$lstVMs.ForeColor = $uiText
$tabControl.BackColor = $uiPanel
$tabControl.ForeColor = $uiText
$progressBar.ForeColor = $uiAccent2
$progressBar.BackColor = $uiInput

Set-ModernButtonIfPresent -Name "btnUnlockSettings" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 12
Set-ModernButtonIfPresent -Name "btnConnect" -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 12
Set-ModernButtonIfPresent -Name "btnCheckTable" -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnRefreshArchive" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnExportCSV" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnExportExcel" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnSaveSQLSettings" -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnClearLogs" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnBrowseLog" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnAddGroup" -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnRemoveGroup" -BackColor $uiDanger -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnTestEmail" -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnSaveSettings" -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnResetSettings" -BackColor $uiWarning -ForeColor $uiButtonText -BorderSize 0 -Radius 10
Set-ModernButtonIfPresent -Name "btnRefresh" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnClearSelection" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnLoadCSV" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-ModernButtonIfPresent -Name "btnDelete" -BackColor $uiDanger -ForeColor $uiButtonText -BorderSize 0 -Radius 12
Set-ModernButtonIfPresent -Name "btnResume" -BackColor $uiAccent -ForeColor $uiButtonText -BorderSize 0 -Radius 12
Set-ModernButtonIfPresent -Name "btnClose" -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
Set-AccessibleControlIfPresent -Name "btnUnlockSettings" -AccessibleName "Unlock settings" -AccessibleDescription "Unlocks the Archive and Settings tabs." -TabIndex 0
Set-AccessibleControlIfPresent -Name "cboVCServer" -AccessibleName "vCenter Server" -AccessibleDescription "Select the vCenter environment to connect to." -TabIndex 1
Set-AccessibleControlIfPresent -Name "btnConnect" -AccessibleName "Connect" -AccessibleDescription "Connect to the selected vCenter server." -TabIndex 2
Set-AccessibleControlIfPresent -Name "txtFilter" -AccessibleName "VM filter" -AccessibleDescription "Filter the virtual machine list." -TabIndex 3
Set-AccessibleControlIfPresent -Name "cboOSFilter" -AccessibleName "OS filter" -AccessibleDescription "Filter virtual machines by operating system." -TabIndex 4
Set-AccessibleControlIfPresent -Name "cboPowerFilter" -AccessibleName "Power state filter" -AccessibleDescription "Filter virtual machines by power state." -TabIndex 5
Set-AccessibleControlIfPresent -Name "lstVMs" -AccessibleName "Virtual machine selection list" -AccessibleDescription "Select one or more VMs or templates for deletion." -TabIndex 6
Set-AccessibleControlIfPresent -Name "btnRefresh" -AccessibleName "Refresh VM list" -AccessibleDescription "Reload the VM list from vCenter." -TabIndex 6
Set-AccessibleControlIfPresent -Name "btnClearSelection" -AccessibleName "Clear VM selection" -AccessibleDescription "Clear the selected VMs." -TabIndex 7
Set-AccessibleControlIfPresent -Name "btnLoadCSV" -AccessibleName "Load VM list from CSV" -AccessibleDescription "Import VM names from a CSV file." -TabIndex 8
Set-AccessibleControlIfPresent -Name "txtReason" -AccessibleName "Deletion reason" -AccessibleDescription "Reason saved to archive and email notification." -TabIndex 9
Set-AccessibleControlIfPresent -Name "chkIncludeTemplates" -AccessibleName "Include templates" -AccessibleDescription "Include VM templates in the selection list." -TabIndex 10
Set-AccessibleControlIfPresent -Name "chkSCCM" -AccessibleName "Delete SCCM record" -TabIndex 11
Set-AccessibleControlIfPresent -Name "chkAD" -AccessibleName "Delete AD record" -TabIndex 12
Set-AccessibleControlIfPresent -Name "chkDNS" -AccessibleName "Delete DNS record" -TabIndex 13
Set-AccessibleControlIfPresent -Name "chkVM" -AccessibleName "Delete virtual machine" -TabIndex 14
Set-AccessibleControlIfPresent -Name "chkArchive" -AccessibleName "Archive deletion result" -TabIndex 15
Set-AccessibleControlIfPresent -Name "chkEmail" -AccessibleName "Send deletion email" -TabIndex 16
Set-AccessibleControlIfPresent -Name "btnResume" -AccessibleName "Resume deletion" -AccessibleDescription "Resume deletion from the last failed VM when progress data exists." -TabIndex 17
Set-AccessibleControlIfPresent -Name "btnDelete" -AccessibleName "Delete selected VMs" -AccessibleDescription "Start the deletion workflow for selected VMs." -TabIndex 18
Set-AccessibleControlIfPresent -Name "btnClose" -AccessibleName "Close" -AccessibleDescription "Close the tool after the workflow is complete." -TabIndex 19

Set-AccessibleControlIfPresent -Name "txtLogs" -AccessibleName "Error details log" -AccessibleDescription "Full error details collected from the tool."
Set-AccessibleControlIfPresent -Name "btnClearLogs" -AccessibleName "Clear logs" -AccessibleDescription "Clear the on-screen LOGS tab text."
Set-AccessibleControlIfPresent -Name "txtSQLServer" -AccessibleName "SQL server"
Set-AccessibleControlIfPresent -Name "txtSQLDatabase" -AccessibleName "SQL database"
Set-AccessibleControlIfPresent -Name "txtSQLTable" -AccessibleName "SQL table"
Set-AccessibleControlIfPresent -Name "chkSQLIntegrated" -AccessibleName "Use Windows authentication for SQL"
Set-AccessibleControlIfPresent -Name "txtSQLUser" -AccessibleName "SQL username"
Set-AccessibleControlIfPresent -Name "txtSQLPass" -AccessibleName "SQL password"
Set-AccessibleControlIfPresent -Name "btnCheckTable" -AccessibleName "Check SQL table"
Set-AccessibleControlIfPresent -Name "dgvArchive" -AccessibleName "Deletion archive data table"
Set-AccessibleControlIfPresent -Name "btnRefreshArchive" -AccessibleName "Refresh archive"
Set-AccessibleControlIfPresent -Name "btnExportCSV" -AccessibleName "Export archive to CSV"
Set-AccessibleControlIfPresent -Name "btnExportExcel" -AccessibleName "Export archive to Excel"
Set-AccessibleControlIfPresent -Name "btnSaveSQLSettings" -AccessibleName "Save SQL settings"
Set-AccessibleControlIfPresent -Name "txtSettingsSCCM" -AccessibleName "SCCM server"
Set-AccessibleControlIfPresent -Name "txtSettingsSiteCode" -AccessibleName "SCCM site code"
Set-AccessibleControlIfPresent -Name "txtSettingsDNS" -AccessibleName "DNS server"
Set-AccessibleControlIfPresent -Name "txtSettingsDNSZone" -AccessibleName "DNS zone"
Set-AccessibleControlIfPresent -Name "txtSettingsDC" -AccessibleName "Domain controller"
Set-AccessibleControlIfPresent -Name "txtSettingsEmail" -AccessibleName "Notification email"
Set-AccessibleControlIfPresent -Name "chkEnableLogging" -AccessibleName "Enable logging"
Set-AccessibleControlIfPresent -Name "chkConfirmDeletion" -AccessibleName "Require deletion confirmation"
Set-AccessibleControlIfPresent -Name "chkDeleteFromDisk" -AccessibleName "Delete VM files from disk permanently"
Set-AccessibleControlIfPresent -Name "numTimeout" -AccessibleName "Power off timeout seconds"
Set-AccessibleControlIfPresent -Name "txtLogPath" -AccessibleName "Log path"
Set-AccessibleControlIfPresent -Name "btnBrowseLog" -AccessibleName "Browse log path"
Set-AccessibleControlIfPresent -Name "chkEnableEmail" -AccessibleName "Enable email notifications"
Set-AccessibleControlIfPresent -Name "txtSMTPServer" -AccessibleName "SMTP server"
Set-AccessibleControlIfPresent -Name "numSMTPPort" -AccessibleName "SMTP port"
Set-AccessibleControlIfPresent -Name "chkSMTPSSL" -AccessibleName "Use SMTP SSL"
Set-AccessibleControlIfPresent -Name "txtSMTPFrom" -AccessibleName "SMTP from address"
Set-AccessibleControlIfPresent -Name "txtNewGroup" -AccessibleName "New email group or address"
Set-AccessibleControlIfPresent -Name "cboGroupType" -AccessibleName "Email group type"
Set-AccessibleControlIfPresent -Name "btnAddGroup" -AccessibleName "Add email group"
Set-AccessibleControlIfPresent -Name "btnRemoveGroup" -AccessibleName "Remove email group"
Set-AccessibleControlIfPresent -Name "btnTestEmail" -AccessibleName "Send test email"
Set-AccessibleControlIfPresent -Name "btnResetSettings" -AccessibleName "Reset settings"
Set-AccessibleControlIfPresent -Name "btnSaveSettings" -AccessibleName "Save settings"
$lblGroupsWindows.ForeColor = $uiAccent
$lblGroupsLinux.ForeColor = $uiWarning
Set-AccessibilityTree -Root $form
#endregion

$script:VMFilterTimer = New-Object System.Windows.Forms.Timer
$script:VMFilterTimer.Interval = 250
$script:PendingVMFilterText = ""
$script:PendingVMOSFilter = "All"
$script:PendingVMPowerFilter = "All"
$script:LastAppliedVMFilterKey = ""
$script:VMFilterTimer.Add_Tick({
    $script:VMFilterTimer.Stop()
    Filter-VMList -FilterText $script:PendingVMFilterText -OSFilter $script:PendingVMOSFilter -PowerFilter $script:PendingVMPowerFilter
})
# ========== Event handler לבחירה ברשימה ==========
$lstVMs.Add_ItemChecked({
    param($sender, $e)
    
    if ($script:IsUpdatingList) { return }
    
    $vmName = $e.Item.Text
    
    if ($e.Item.Checked) {
        if ($script:SelectedVMNames -notcontains $vmName) {
            [void]$script:SelectedVMNames.Add($vmName)
        }
    }
    else {
        $script:SelectedVMNames.Remove($vmName)
    }
    
    $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"
    
})


# ========== Email Groups Event Handlers ==========

# משתנה לעקוב אחרי הרשימה הנבחרת
$script:selectedGroupList = $null

# Click Events - לזהות איזו רשימה נבחרה
$script:lstGroupsCommon.Add_MouseDown({
    $script:selectedGroupList = "Other"
    $script:lstGroupsWindows.ClearSelected()
    $script:lstGroupsLinux.ClearSelected()
})

$script:lstGroupsWindows.Add_MouseDown({
    $script:selectedGroupList = "Windows"
    $script:lstGroupsCommon.ClearSelected()
    $script:lstGroupsLinux.ClearSelected()
})

$script:lstGroupsLinux.Add_MouseDown({
    $script:selectedGroupList = "Linux"
    $script:lstGroupsCommon.ClearSelected()
    $script:lstGroupsWindows.ClearSelected()
})

# Add Group Button
$btnAddGroup.Add_Click({
    $groupName = $txtNewGroup.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($groupName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a group name.",
            "Validation",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $txtNewGroup.Focus()
        return
    }
    
    # בחר את הרשימה לפי ה-ComboBox
    $selectedType = $cboGroupType.SelectedItem.ToString()
    $targetList = $null
    
    switch ($selectedType) {
        "Other"   { $targetList = $script:lstGroupsCommon }
        "Windows" { $targetList = $script:lstGroupsWindows }
        "Linux"   { $targetList = $script:lstGroupsLinux }
    }
    
    if ($null -eq $targetList) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please select a group type.",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }
    
    # בדיקה אם הקבוצה כבר קיימת
    $exists = $false
    for ($i = 0; $i -lt $targetList.Items.Count; $i++) {
        if ($targetList.Items[$i].ToString() -eq $groupName) {
            $exists = $true
            break
        }
    }
    
    if ($exists) {
        [System.Windows.Forms.MessageBox]::Show(
            "Group '$groupName' already exists in $selectedType list.",
            "Duplicate",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    # הוסף את הקבוצה
    [void]$targetList.Items.Add($groupName)
    $txtNewGroup.Text = ""
    $txtNewGroup.Focus()
    
    Write-Log "Added group '$groupName' to $selectedType list"
})

# Remove Group Button
$btnRemoveGroup.Add_Click({
    # מצא איזו רשימה יש בה פריט נבחר
    $targetList = $null
    $listName = ""
    
    if ($script:lstGroupsCommon.SelectedIndex -ge 0) {
        $targetList = $script:lstGroupsCommon
        $listName = "Other"
    }
    elseif ($script:lstGroupsWindows.SelectedIndex -ge 0) {
        $targetList = $script:lstGroupsWindows
        $listName = "Windows"
    }
    elseif ($script:lstGroupsLinux.SelectedIndex -ge 0) {
        $targetList = $script:lstGroupsLinux
        $listName = "Linux"
    }
    
    if ($null -eq $targetList -or $targetList.SelectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please click on a group in one of the lists to select it, then click Remove.",
            "No Selection",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return
    }
    
    $selectedIndex = $targetList.SelectedIndex
    $groupName = $targetList.Items[$selectedIndex].ToString()
    
    $confirm = [System.Windows.Forms.MessageBox]::Show(
        "Remove group '$groupName' from $listName list?",
        "Confirm Remove",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $targetList.Items.RemoveAt($selectedIndex)
        Write-Log "Removed group '$groupName' from $listName list"
    }
})

# Test Email Button
$btnTestEmail.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtSMTPServer.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please configure SMTP server first.",
            "SMTP Not Configured",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if ([string]::IsNullOrWhiteSpace($txtSettingsEmail.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please configure Notification Email in AD Settings first.",
            "Email Not Configured",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    $lblSettingsStatus.Text = "Sending test email..."
    $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
$testBody = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            padding: 0; 
            margin: 0; 
            background-color: #e9ecef; 
        }
        .container { 
            max-width: 500px; 
            margin: 30px auto; 
            background-color: #ffffff; 
            border-radius: 10px; 
            box-shadow: 0 4px 15px rgba(0,0,0,0.1); 
            overflow: hidden; 
        }
        .header { 
            background: linear-gradient(135deg, #0066cc 0%, #004499 100%); 
            color: white; 
            padding: 25px; 
            text-align: center; 
        }
        .header h2 { 
            margin: 0; 
            font-size: 22px; 
            font-weight: 600; 
        }
        .icon { 
            font-size: 40px; 
            margin-bottom: 10px; 
        }
        .content { 
            padding: 30px; 
            text-align: center; 
            color: #333333; 
        }
        .success-badge { 
            background-color: #d4edda; 
            color: #155724; 
            padding: 12px 20px; 
            border-radius: 25px; 
            display: inline-block; 
            margin-bottom: 15px; 
            font-weight: 500; 
        }
        .message { 
            color: #666666; 
            font-size: 14px; 
            margin-bottom: 25px; 
        }
        .details { 
            background-color: #f8f9fa; 
            border-radius: 8px; 
            padding: 20px; 
            text-align: left; 
        }
        .detail-row { 
            display: flex; 
            justify-content: space-between; 
            padding: 8px 0; 
            border-bottom: 1px solid #e9ecef; 
        }
        .detail-row:last-child { 
            border-bottom: none; 
        }
        .detail-label { 
            color: #888888; 
            font-size: 13px; 
        }
        .detail-value { 
            color: #333333; 
            font-weight: 500; 
            font-size: 13px; 
        }
        .footer { 
            background-color: #f8f9fa; 
            padding: 15px; 
            text-align: center; 
            font-size: 11px; 
            color: #999999; 
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="icon"></div>
            <h2>VM Deletion Tool</h2>
        </div>
        <div class="content">
            <div class="success-badge">Test Email</div>
            <p class="message">This is a test email from <strong>VM Deletion Tool</strong>.<br>If you received this email, your SMTP settings are configured correctly.</p>
            <div class="details">
                <div class="detail-row">
                    <span class="detail-label">Sent at</span>
                    <span class="detail-value">$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</span>
                </div>
                <div class="detail-row">
                    <span class="detail-label">Sent by</span>
                    <span class="detail-value">$env:USERDOMAIN\$env:USERNAME</span>
                </div>
            </div>
        </div>
        <div class="footer">
            VM Deletion Tool - IT Infrastructure Management
        </div>
    </div>
</body>
</html>
"@
        
        $mailParams = @{
            From       = $txtSMTPFrom.Text
            To         = $txtSettingsEmail.Text
            Subject    = "[VM Deletion Tool] Test Email - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            Body       = $testBody
            BodyAsHtml = $true
            SmtpServer = $txtSMTPServer.Text
            Port       = [int]$numSMTPPort.Value
            Priority   = "Normal"
        }
        
        if ($chkSMTPSSL.Checked) {
            $mailParams.UseSsl = $true
        }
        
        Send-MailMessage @mailParams -ErrorAction Stop
        
        $lblSettingsStatus.Text = "OK - Test email sent successfully to $($txtSettingsEmail.Text)"
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        
        [System.Windows.Forms.MessageBox]::Show(
            "Test email sent successfully!`n`nCheck inbox: $($txtSettingsEmail.Text)",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    catch {
        $errorMessage = "Failed to send test email: $($_.Exception.Message)"
        $lblSettingsStatus.Text = "Failed - Email failed. See LOGS."
        $lblSettingsStatus.ForeColor = $uiDanger
        Write-Log $errorMessage -Level "ERROR"
    }
})


$btnUnlockSettings.Add_Click({
    # If tabs are currently enabled, lock them
    if ($tabArchive.Enabled) {
        $tabArchive.Enabled = $false
        $tabSettings.Enabled = $false
        $btnUnlockSettings.Text = "🔒"
        $btnUnlockSettings.ForeColor = [System.Drawing.Color]::Black
    }
    else {
        # If tabs are disabled, show password prompt
        $passwordForm = New-Object System.Windows.Forms.Form
        $passwordForm.Text = "Unlock Settings"
        $passwordForm.Size = New-Object System.Drawing.Size(250, 150)
        $passwordForm.StartPosition = "CenterScreen"
        $passwordForm.FormBorderStyle = "FixedDialog"
        $passwordForm.MaximizeBox = $false
        $passwordForm.TopMost = $true

        $lblPassword = New-Object System.Windows.Forms.Label
        $lblPassword.Location = New-Object System.Drawing.Point(20, 20)
        $lblPassword.Size = New-Object System.Drawing.Size(200, 20)
        $lblPassword.Text = "Enter Settings Password:"
        $passwordForm.Controls.Add($lblPassword)

        $txtPasswordInput = New-Object System.Windows.Forms.TextBox
        $txtPasswordInput.Location = New-Object System.Drawing.Point(20, 50)
        $txtPasswordInput.Size = New-Object System.Drawing.Size(200, 23)
        $txtPasswordInput.UseSystemPasswordChar = $true
        $passwordForm.Controls.Add($txtPasswordInput)

        $btnSubmit = New-Object System.Windows.Forms.Button
        $btnSubmit.Location = New-Object System.Drawing.Point(20, 80)
        $btnSubmit.Size = New-Object System.Drawing.Size(90, 30)
        $btnSubmit.Text = "Unlock"
        $passwordForm.Controls.Add($btnSubmit)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(130, 80)
        $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
        $btnCancel.Text = "Cancel"
        $passwordForm.Controls.Add($btnCancel)

        $btnSubmit.Add_Click({
            if ($txtPasswordInput.Text -eq $script:settingsPassword) {
                $tabArchive.Enabled = $true
                $tabSettings.Enabled = $true
                $btnUnlockSettings.Text = "🔓"
                $btnUnlockSettings.ForeColor = [System.Drawing.Color]::Green
                $passwordForm.Close()
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Incorrect password", 
                    "Access Denied", 
                    [System.Windows.Forms.MessageBoxButtons]::OK, 
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
        })

        $btnCancel.Add_Click({ $passwordForm.Close() })
        $passwordForm.AcceptButton = $btnSubmit
        $passwordForm.CancelButton = $btnCancel

        $passwordForm.ShowDialog()
    }
})

$btnUnlockSettings.Add_DoubleClick({
    $tabArchive.Enabled = $false
    $tabSettings.Enabled = $false
    $btnUnlockSettings.Text = "🔒"
    $btnUnlockSettings.ForeColor = [System.Drawing.Color]::Black

    [System.Windows.Forms.MessageBox]::Show(
        "Settings and Archive tabs have been locked.", 
        "Locked", 
        [System.Windows.Forms.MessageBoxButtons]::OK, 
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
})

# ========== vCenter Connect Button ==========
$btnConnect.Add_Click({
    if ($script:vcConnected) {
        # Disconnect
        try {
            Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue
            $script:vcConnected = $false
            $btnConnect.Text = "Connect"
            $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
            $lblConnectionStatus.Text = "Disconnected"
            $lblConnectionStatus.ForeColor = $uiMuted
            $lstVMs.Items.Clear()
            $script:allVMs = @()
            $script:allVMsData = @()
            $lblVMCount.Text = "0 VMs"
            $txtFilter.Enabled = $false
            $cboOSFilter.Enabled = $false
            $cboPowerFilter.Enabled = $false
            $btnRefresh.Enabled = $false
            $btnClearSelection.Enabled = $false
            $btnDelete.Enabled = $false
            $lstVMs.Enabled = $true
            $lstVMs.BackColor = $uiInput
            $lstVMs.ForeColor = $uiText
            $txtReason.Enabled = $false
            Write-Log "Disconnected from vCenter"
        } catch {
            Write-Log "Error disconnecting: $($_.Exception.Message)" -Level "ERROR"
        }
        return
    }
    
    # Connect
    $selectedIndex = $cboVCServer.SelectedIndex
    if ($selectedIndex -lt 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a vCenter server.", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $vcServer = $script:vCenterServers[$selectedIndex].Server
    $vcName = $script:vCenterServers[$selectedIndex].Name
    
    $lblConnectionStatus.Text = "Connecting to $vcName..."
    $lblConnectionStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $btnConnect.Enabled = $false
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        # Check if PowerCLI is loaded
        if (-not (Get-Module -Name VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) {
            Import-Module VMware.VimAutomation.Core -ErrorAction Stop
        }
        
        # Set PowerCLI configuration
        Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        
        # Connect to vCenter
        Connect-VIServer -Server $vcServer -ErrorAction Stop | Out-Null
        
        $script:vcConnected = $true
        $btnConnect.Text = "Disconnect"
        $btnConnect.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
        $lblConnectionStatus.Text = "Connected to $vcName ($vcServer)"
        $lblConnectionStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        
        # Enable controls
        $txtFilter.Enabled = $true
        $cboOSFilter.Enabled = $true
        $cboPowerFilter.Enabled = $true
        $btnRefresh.Enabled = $true
        $btnClearSelection.Enabled = $true
        $btnLoadCSV.Enabled = $true
        $btnResume.Enabled = $true
        $btnDelete.Enabled = $true
        $lstVMs.Enabled = $true
        $txtReason.Enabled = $true
        $chkIncludeTemplates.Enabled = $true
        
        # Load VMs
        Get-VMList
        
        Write-Log "Connected to vCenter: $vcServer"
    }
    catch {
        $errorMessage = "vCenter connection failed: $($_.Exception.Message)"
        $lblConnectionStatus.Text = "Failed - vCenter failed. See LOGS."
        $lblConnectionStatus.ForeColor = $uiDanger
        if ($tooltip) {
            $tooltip.SetToolTip($lblConnectionStatus, "Failed - vCenter failed. See LOGS.")
        }
        Write-Log $errorMessage -Level "ERROR"
    }
    finally {
        $btnConnect.Enabled = $true
    }
})

# ========== Filter TextBox ==========
$txtFilter.Add_TextChanged({
    Schedule-VMFilter
})

# ========== OS Filter ComboBox ==========
$cboOSFilter.Add_SelectedIndexChanged({
    Schedule-VMFilter
})

# ========== Power Filter ComboBox ==========
$cboPowerFilter.Add_SelectedIndexChanged({
    Schedule-VMFilter
})

# ========== Include Templates Checkbox ==========
$chkIncludeTemplates.Add_CheckedChanged({
    $script:IncludeTemplates = $chkIncludeTemplates.Checked
    if ($script:vcConnected) {
        Get-VMList
    }
})

# ========== Refresh Button ==========
$btnRefresh.Add_Click({
    $lblConnectionStatus.Text = "Refreshing VM list..."
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    Get-VMList
    $lblConnectionStatus.Text = "VM list refreshed"
})
# ========== CSV List Button ==========
$btnLoadCSV.Add_Click({   
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select CSV file with VM list"

    if ($openFileDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return
    }

    try {
        $csvData = Import-Csv -Path $openFileDialog.FileName

        $hasVMName      = $csvData | Get-Member -Name "VMName"      -MemberType NoteProperty
        $hasVCenter     = $csvData | Get-Member -Name "vCenter"     -MemberType NoteProperty
        $hasEnvironment = $csvData | Get-Member -Name "Environment" -MemberType NoteProperty

        if (-not $hasVMName) {
            [System.Windows.Forms.MessageBox]::Show(
                "CSV file must contain a column named 'VMName'.",
                "Invalid CSV Format",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            return
        }

        # ===== Load all VM names once (fast) =====
        $vmHash = [System.Collections.Generic.HashSet[string]]::new()
        if ($script:vcConnected) {
            foreach ($vmData in $script:allVMsData) {
                $vmHash.Add($vmData.Name) | Out-Null
            }
        }

        # ===== Preview Form =====
        $previewForm = New-Object System.Windows.Forms.Form
        $previewForm.Text = "CSV Preview - Select VMs to Import"
        $previewForm.Size = New-Object System.Drawing.Size(760, 520)
        $previewForm.StartPosition = "CenterScreen"
        $previewForm.FormBorderStyle = "FixedDialog"
        $previewForm.MaximizeBox = $false
        $previewForm.TopMost = $true
        $previewForm.BackColor = $uiPanel

        # ===== DataGrid =====
        $dataGrid = New-Object System.Windows.Forms.DataGridView
        $dataGrid.Location = New-Object System.Drawing.Point(10, 10)
        $dataGrid.Size = New-Object System.Drawing.Size(720, 380)
        $dataGrid.AllowUserToAddRows = $false
        $dataGrid.AllowUserToDeleteRows = $false
        $dataGrid.RowHeadersVisible = $false
        $dataGrid.SelectionMode = "FullRowSelect"
        $dataGrid.MultiSelect = $false
        $dataGrid.AutoSizeColumnsMode = "Fill"
        $dataGrid.AllowUserToResizeRows = $false
        $dataGrid.BackgroundColor = $uiInput
        $dataGrid.GridColor = $uiBorderSoft
        $dataGrid.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
        $dataGrid.EnableHeadersVisualStyles = $false
        $dataGrid.DefaultCellStyle.BackColor = $uiInput
        $dataGrid.DefaultCellStyle.ForeColor = $uiText
        $dataGrid.DefaultCellStyle.SelectionBackColor = $uiAccent2
        $dataGrid.DefaultCellStyle.SelectionForeColor = $uiButtonText
        $dataGrid.ColumnHeadersDefaultCellStyle.BackColor = $uiPanel2
        $dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = $uiText
        $previewForm.Controls.Add($dataGrid)

        # ===== Columns =====
        $checkCol = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
        $checkCol.Name = "Select"
        $checkCol.HeaderText = "Select"
        $checkCol.Width = 40
        $checkCol.AutoSizeMode = "None"
        $dataGrid.Columns.Add($checkCol) | Out-Null

        $statusCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $statusCol.Name = "Status"
        $statusCol.HeaderText = "Status"
        $statusCol.Width = 90
        $statusCol.ReadOnly = $true
        $statusCol.AutoSizeMode = "None"
        $dataGrid.Columns.Add($statusCol) | Out-Null

        if ($hasVCenter) { 
            $dataGrid.Columns.Add("vCenter", "vCenter") | Out-Null 
        }
        if ($hasEnvironment) {
            $dataGrid.Columns.Add("Environment", "Environment") | Out-Null
        }
        $dataGrid.Columns.Add("VMName", "VM Name") | Out-Null

        # ===== Populate Rows =====
        $foundCount = 0
        $notFoundCount = 0
        
        foreach ($row in $csvData) {
            $vmName = $row.VMName.Trim()
            if ([string]::IsNullOrWhiteSpace($vmName)) { continue }

            $status = "Unknown"
            $exists = $false

            if ($script:vcConnected) {
                if ($vmHash.Contains($vmName)) {
                    $status = "Found"
                    $exists = $true
                    $foundCount++
                }
                else {
                    $status = "Not Found"
                    $notFoundCount++
                }
            }

            $rowData = @($exists, $status)  # רק Found מסומן כברירת מחדל
            if ($hasVCenter) { $rowData += $row.vCenter }
            if ($hasEnvironment) { $rowData += $row.Environment }
            $rowData += $vmName

            $index = $dataGrid.Rows.Add($rowData)

            # Coloring
            if ($script:vcConnected) {
                if ($exists) {
                    $dataGrid.Rows[$index].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(39, 72, 58)
                    $dataGrid.Rows[$index].DefaultCellStyle.ForeColor = $uiText
                }
                else {
                    $dataGrid.Rows[$index].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(82, 48, 55)
                    $dataGrid.Rows[$index].DefaultCellStyle.ForeColor = $uiMuted
                }
            }
        }

        # ===== Summary =====
        $lblSummary = New-Object System.Windows.Forms.Label
        $lblSummary.Location = New-Object System.Drawing.Point(10, 400)
        $lblSummary.Size = New-Object System.Drawing.Size(700, 20)
        $lblSummary.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        if ($script:vcConnected) {
            $lblSummary.Text = "Total: $($dataGrid.Rows.Count) | Found: $foundCount | Not Found: $notFoundCount"
        }
        else {
            $lblSummary.Text = "Total VMs in CSV: $($dataGrid.Rows.Count) (Connect to vCenter to verify)"
        }
        $lblSummary.ForeColor = $uiMuted
        $previewForm.Controls.Add($lblSummary)

        # ===== Select All Button =====
        $btnSelectAll = New-Object System.Windows.Forms.Button
        $btnSelectAll.Location = New-Object System.Drawing.Point(10, 430)
        $btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
        $btnSelectAll.Text = "Select All"
        $btnSelectAll.FlatStyle = "Flat"
        Set-ModernButton -Button $btnSelectAll -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
        $btnSelectAll.Add_Click({
            foreach ($r in $dataGrid.Rows) {
                $r.Cells["Select"].Value = $true
            }
        })
        $previewForm.Controls.Add($btnSelectAll)

        # ===== Select None Button =====
        $btnSelectNone = New-Object System.Windows.Forms.Button
        $btnSelectNone.Location = New-Object System.Drawing.Point(120, 430)
        $btnSelectNone.Size = New-Object System.Drawing.Size(100, 30)
        $btnSelectNone.Text = "Select None"
        $btnSelectNone.FlatStyle = "Flat"
        Set-ModernButton -Button $btnSelectNone -BackColor $uiPanel2 -ForeColor $uiButtonText -BorderSize 1 -Radius 10
        $btnSelectNone.Add_Click({
            foreach ($r in $dataGrid.Rows) {
                $r.Cells["Select"].Value = $false
            }
        })
        $previewForm.Controls.Add($btnSelectNone)

        # ===== Select Found Only Button =====
        $btnSelectFound = New-Object System.Windows.Forms.Button
        $btnSelectFound.Location = New-Object System.Drawing.Point(230, 430)
        $btnSelectFound.Size = New-Object System.Drawing.Size(140, 30)
        $btnSelectFound.Text = "Select Found Only"
        $btnSelectFound.FlatStyle = "Flat"
        Set-ModernButton -Button $btnSelectFound -BackColor $uiSuccess -ForeColor $uiButtonText -BorderSize 0 -Radius 10
        $btnSelectFound.Add_Click({
            foreach ($r in $dataGrid.Rows) {
                $r.Cells["Select"].Value = ($r.Cells["Status"].Value -eq "Found")
            }
        })
        $previewForm.Controls.Add($btnSelectFound)

        # ===== Import Button =====
        $btnImport = New-Object System.Windows.Forms.Button
        $btnImport.Location = New-Object System.Drawing.Point(520, 430)
        $btnImport.Size = New-Object System.Drawing.Size(100, 30)
        $btnImport.Text = "Import"
        $btnImport.FlatStyle = "Flat"
        Set-ModernButton -Button $btnImport -BackColor $uiAccent2 -ForeColor $uiButtonText -BorderSize 0 -Radius 10
        $btnImport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $btnImport.Add_Click({
            $script:SelectedVMNames.Clear()
            $script:CSVImportData = @()
            $importedCount = 0
            $skippedCount = 0

            foreach ($row in $dataGrid.Rows) {
                if ($row.Cells["Select"].Value -eq $true) {
                    $vmName = $row.Cells["VMName"].Value
                    if ([string]::IsNullOrWhiteSpace($vmName)) { continue }

                    # דלג על VMs שלא נמצאו
                    if ($script:vcConnected -and $row.Cells["Status"].Value -eq "Not Found") {
                        $skippedCount++
                        continue
                    }

                    $vmData = @{ VMName = $vmName }
                    if ($hasVCenter) {
                        $vmData.vCenter = $row.Cells["vCenter"].Value
                    }
                    if ($hasEnvironment) {
                        $vmData.Environment = $row.Cells["Environment"].Value
                    }

                    $script:CSVImportData += $vmData
                    $script:SelectedVMNames.Add($vmName) | Out-Null
                    $importedCount++
                }
            }

            $previewForm.Tag = @{
                Imported = $importedCount
                Skipped  = $skippedCount
            }
            $previewForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $previewForm.Close()
        })
        $previewForm.Controls.Add($btnImport)

        # ===== Cancel =====
        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Location = New-Object System.Drawing.Point(630, 430)
        $btnCancel.Size = New-Object System.Drawing.Size(100, 30)
        $btnCancel.Text = "Cancel"
        $btnCancel.FlatStyle = "Flat"
        $btnCancel.Add_Click({
            $previewForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $previewForm.Close()
        })
        $previewForm.Controls.Add($btnCancel)

        # ===== Show Dialog =====
        $result = $previewForm.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            # עדכן CheckBoxes ברשימה הראשית
            $script:IsUpdatingList = $true
            foreach ($item in $lstVMs.Items) {
                $item.Checked = ($script:SelectedVMNames -contains $item.Text)
            }
            $script:IsUpdatingList = $false

            $lblSelectedCount.Text = "$($script:SelectedVMNames.Count) selected"

            # הודעה למשתמש
            $importInfo = $previewForm.Tag
            $message = "Imported $($importInfo.Imported) VMs."
            if ($importInfo.Skipped -gt 0) {
                $message += "`n`n$($importInfo.Skipped) VMs were not found and skipped."
            }

            [System.Windows.Forms.MessageBox]::Show(
                $message,
                "Import Complete",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    catch {
        Write-Log "Error loading CSV: $($_.Exception.Message)" -Level "ERROR"
    }
})
# ========== Clear Selection Button ==========
$btnClearSelection.Add_Click({
    $script:SelectedVMNames.Clear()
    foreach ($item in $lstVMs.Items) {
        $item.Checked = $false
    }
    $lblSelectedCount.Text = "0 selected"

})

# ========== SQL Integrated Auth Checkbox ==========
$chkSQLIntegrated.Add_CheckedChanged({
    $enabled = -not $chkSQLIntegrated.Checked
    $txtSQLUser.Enabled = $enabled
    $txtSQLPass.Enabled = $enabled
    $lblSQLUser.Enabled = $enabled
    $lblSQLPass.Enabled = $enabled
})

# ========== Check SQL Table Button ==========
$btnCheckTable.Add_Click({
    # Update settings from UI
    $script:DefaultSettings.SQLServer = $txtSQLServer.Text
    $script:DefaultSettings.SQLDatabase = $txtSQLDatabase.Text
    $script:DefaultSettings.SQLTable = $txtSQLTable.Text
    $script:DefaultSettings.SQLIntegratedAuth = $chkSQLIntegrated.Checked
    $script:DefaultSettings.SQLUsername = $txtSQLUser.Text
    $script:DefaultSettings.SQLPassword = $txtSQLPass.Text
    
    $lblArchiveStatus.Text = "Checking SQL connection..."
    $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
    
    $connectionString = Get-SQLConnectionString
    
    if ([string]::IsNullOrWhiteSpace($connectionString)) {
        $lblArchiveStatus.Text = "Failed - Invalid SQL configuration"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
        return
    }
    
    # Test connection
    $connTest = Test-SQLConnection -ConnectionString $connectionString
    
    if (-not $connTest.Success) {
        $lblArchiveStatus.Text = "Failed - SQL connection failed. See LOGS."
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "SQL connection test failed: $($connTest.Message)" -Level "ERROR"
        return
    }
    
    # Check table exists
    $tableCheck = Test-SQLTableExists -ConnectionString $connectionString -TableName $($script:DefaultSettings.SQLTable)

    
    if (-not $tableCheck.Success) {
        $lblArchiveStatus.Text = "Failed - SQL table check failed. See LOGS."
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "SQL table check failed: $($tableCheck.Message)" -Level "ERROR"
        return
    }
    
    if ($tableCheck.Exists) {
        $lblArchiveStatus.Text = "OK - Table '$($script:DefaultSettings.SQLTable)' exists with $($tableCheck.Columns.Count) columns"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        $script:sqlTableExists = $true
        $btnRefreshArchive.Enabled = $true
        $btnExportCSV.Enabled = $true
        $btnExportExcel.Enabled = $true
        
        # Load archive data
        $btnRefreshArchive.PerformClick()
    } else {
        $createTable = [System.Windows.Forms.MessageBox]::Show(
            "Table '$($script:DefaultSettings.SQLTable)' does not exist.`n`nDo you want to create it?",
            "Create Table",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($createTable -eq [System.Windows.Forms.DialogResult]::Yes) {
            $createResult = New-SQLArchiveTable -ConnectionString $connectionString -TableName $($script:DefaultSettings.SQLTable)
            
            if ($createResult.Success) {
                $lblArchiveStatus.Text = "OK - $($createResult.Message)"
                $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
                $script:sqlTableExists = $true
                $btnRefreshArchive.Enabled = $true
                $btnExportCSV.Enabled = $true
                $btnExportExcel.Enabled = $true
            } else {
                $lblArchiveStatus.Text = "Failed - SQL table create failed. See LOGS."
                $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
                Write-Log "SQL table create failed: $($createResult.Message)" -Level "ERROR"
            }
        } else {
            $lblArchiveStatus.Text = "Table does not exist - archive disabled"
            $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(255, 193, 7)
        }
    }
})

# ========== Refresh Archive Button ==========
$btnRefreshArchive.Add_Click({
    $connectionString = Get-SQLConnectionString
   
      # בדוק חיבור
    try {
        $testConn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        $testConn.Open()
        $testCmd = $testConn.CreateCommand()
        $testCmd.CommandText = "SELECT @@SERVERNAME as Srv, DB_NAME() as DB"
        $testReader = $testCmd.ExecuteReader()
        if ($testReader.Read()) {
            Write-Log "📊 ARCHIVE TAB Connected to: Server='$($testReader['Srv'])', DB='$($testReader['DB'])'" -Level "INFO"
        }
        $testReader.Close()
        
        # ספור רשומות
        $countCmd = $testConn.CreateCommand()
        $countCmd.CommandText = "SELECT COUNT(*) FROM [$($script:DefaultSettings.SQLTable)]"
        $count = $countCmd.ExecuteScalar()
        Write-Log "📊 Records in [$($script:DefaultSettings.SQLTable)]: $count" -Level "INFO"
        
        $testConn.Close()
    }
    catch {
        Write-Log "❌ Archive test connection failed: $($_.Exception.Message)" -Level "ERROR"
    }
    
    $archiveResult = Get-ArchiveData -ConnectionString $connectionString -TableName $script:DefaultSettings.SQLTable
    
    if ($archiveResult.Success) {
        $dgvArchive.DataSource = $archiveResult.Data
        $lblRecordCount.Text = "$($archiveResult.Data.Rows.Count) records"
        Write-Log "✅ Archive loaded: $($archiveResult.Data.Rows.Count) records" -Level "INFO"
        
        # Hide Id column if exists
        if ($dgvArchive.Columns["Id"]) {
            $dgvArchive.Columns["Id"].Visible = $false
        }
    } else {
        Write-Log "Failed to load archive: $($archiveResult.Message)" -Level "ERROR"
    }
})

# ========== Export CSV Button ==========
$btnExportCSV.Add_Click({
    if ($dgvArchive.DataSource -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "VMDeletionArchive_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportResult = Export-ArchiveToFile -Data $dgvArchive.DataSource -FilePath $saveDialog.FileName -Format "CSV"
        if ($exportResult.Success) {
            [System.Windows.Forms.MessageBox]::Show($exportResult.Message, "Export", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            Write-Log "Failed to export archive to CSV: $($exportResult.Message)" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show("Failed - Export failed. See LOGS.", "Export", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# ========== Export Excel Button ==========
$btnExportExcel.Add_Click({
    if ($dgvArchive.DataSource -eq $null -or $dgvArchive.DataSource.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv"
    $saveDialog.FileName = "VMDeletionArchive_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $exportResult = Export-ArchiveToExcel -Data $dgvArchive.DataSource -FilePath $saveDialog.FileName
        if ($exportResult.Success) {
            [System.Windows.Forms.MessageBox]::Show($exportResult.Message, "Export", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        else {
            Write-Log "Failed to export archive to Excel: $($exportResult.Message)" -Level "ERROR"
            [System.Windows.Forms.MessageBox]::Show("Failed - Export failed. See LOGS.", "Export", 
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})

# ========== Save SQL Settings Button ==========
$btnSaveSQLSettings.Add_Click({
    $script:DefaultSettings.SQLServer = $txtSQLServer.Text
    $script:DefaultSettings.SQLDatabase = $txtSQLDatabase.Text
    $script:DefaultSettings.SQLTable = $txtSQLTable.Text
    $script:DefaultSettings.SQLIntegratedAuth = $chkSQLIntegrated.Checked
    $script:DefaultSettings.SQLUsername = $txtSQLUser.Text
    $script:DefaultSettings.SQLPassword = $txtSQLPass.Text
    
    if (Save-Settings) {
        $lblArchiveStatus.Text = "OK - SQL settings saved"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
    } else {
        $lblArchiveStatus.Text = "Failed - Failed to save settings"
        $lblArchiveStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

# ========== Browse Log Path Button ==========
$btnBrowseLog.Add_Click({
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select Log Folder"
    $folderDialog.SelectedPath = $txtLogPath.Text
    
    if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtLogPath.Text = $folderDialog.SelectedPath
    }
})

# ========== Save Settings Button ==========
$btnSaveSettings.Add_Click({
    # Update all settings from UI
    $script:DefaultSettings.SCCMServer = $txtSettingsSCCM.Text
    $script:DefaultSettings.SCCMSiteCode = $txtSettingsSiteCode.Text
    $script:DefaultSettings.DNSServer = $txtSettingsDNS.Text
    $script:DefaultSettings.DNSZone = $txtSettingsDNSZone.Text
    $script:DefaultSettings.DomainController = $txtSettingsDC.Text
    $script:DefaultSettings.EmailNotification = $txtSettingsEmail.Text
    $script:DefaultSettings.EnableLogging = $chkEnableLogging.Checked
    $script:DefaultSettings.ConfirmDeletion = $chkConfirmDeletion.Checked
    $script:DefaultSettings.DeleteFromDisk = $chkDeleteFromDisk.Checked
    $script:DefaultSettings.PowerOffTimeout = [int]$numTimeout.Value
    $script:DefaultSettings.LogPath = $txtLogPath.Text
    
    # Email settings
    $script:DefaultSettings.EnableEmailNotification = $chkEnableEmail.Checked
    $script:DefaultSettings.SMTPServer = $txtSMTPServer.Text
    $script:DefaultSettings.SMTPPort = [int]$numSMTPPort.Value
    $script:DefaultSettings.SMTPUseSSL = $chkSMTPSSL.Checked
    $script:DefaultSettings.SMTPFrom = $txtSMTPFrom.Text
    
    # Save email groups
    Save-EmailGroups
    
    if (Save-Settings) {
        $lblSettingsStatus.Text = "OK - Settings saved successfully at $(Get-Date -Format 'HH:mm:ss')"
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
        Write-Log "Settings saved"
    } else {
        $lblSettingsStatus.Text = "Failed - Failed to save settings"
        $lblSettingsStatus.ForeColor = [System.Drawing.Color]::Red
    }
    if (Save-LocalConfig) {
        [System.Windows.Forms.MessageBox]::Show(
            "Settings saved locally.",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    else {
        Write-Log "Failed to save settings locally." -Level "ERROR"
    }
})

# ========== Reset Settings Button ==========
$btnResetSettings.Add_Click({
    if (Reset-Configuration) {
        Refresh-SettingsForm
        [System.Windows.Forms.MessageBox]::Show(
            "Settings have been reset successfully!",
            "Settings Reset",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# ========== Delete VM Button ==========
$btnDelete.Add_Click({
    # בדיקה לפי SelectedVMNames
    if ($script:SelectedVMNames.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one VM to delete.", "No Selection", 
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # בניית רשימת VMs מ-SelectedVMNames
    $selectedVMs = @()
    $selectedVCIndex = $cboVCServer.SelectedIndex
    $defaultEnvironment = if ($selectedVCIndex -ge 0) {
        $script:vCenterServers[$selectedVCIndex].Name
    }
    else {
        $script:RunEnvironment
    }

    foreach ($vmName in $script:SelectedVMNames) {
        $vmData = $script:allVMsData | Where-Object { $_.Name -eq $vmName } | Select-Object -First 1
        $csvVMData = $script:CSVImportData | Where-Object { $_.VMName -eq $vmName } | Select-Object -First 1
        $osType = if ($vmData) { $vmData.OSType } else { "Unknown" }
        $vmEnvironment = if ($csvVMData -and $csvVMData.Environment) {
            $csvVMData.Environment
        }
        elseif ($csvVMData -and $csvVMData.vCenter) {
            Resolve-VCenterEnvironment -Value $csvVMData.vCenter
        }
        else {
            $defaultEnvironment
        }
        
        $selectedVMs += @{
            Name        = $vmName
            OSType      = $osType
            Environment = $vmEnvironment
        }
    }
    
    $vmNames = ($selectedVMs | ForEach-Object { $_.Name }) -join ", "
    $reason = $txtReason.Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($reason)) {
        $reason = "No reason provided"
    }

# Confirmation
if ($script:DefaultSettings.ConfirmDeletion) {
    # ספור Templates בבחירה
    $templateCount = 0
    $vmOnlyCount = 0
    $vmNamesList = @()
    
    foreach ($vmInfo in $selectedVMs) {
        $vmData = $script:allVMsData | Where-Object { $_.Name -eq $vmInfo.Name } | Select-Object -First 1
        if ($vmData -and $vmData.IsTemplate) {
            $templateCount++
        } else {
            $vmOnlyCount++
        }
        # הוספת שם ה-VM לרשימה
        $vmNamesList += $vmInfo.Name
    }
    
    # בניית מחרוזת עם כל השמות
    $vmNamesText = $vmNamesList -join "`n"
    
    $itemType = if ($templateCount -gt 0 -and $vmOnlyCount -gt 0) {
        "$vmOnlyCount VM(s) and $templateCount Template(s)"
    } elseif ($templateCount -gt 0) {
        "$templateCount Template(s)"
    } else {
        "$($selectedVMs.Count) VM(s)"
    }
    
    # יצירת חלון WPF מודרני
    Add-Type -AssemblyName PresentationFramework
    
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Confirm Deletion" 
        Height="400" 
        Width="300"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        MinHeight="400"
        MinWidth="300"
        Background="#1E1E26"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True"
        TextOptions.TextFormattingMode="Display"
        TextOptions.TextRenderingMode="ClearType"
        Topmost="True">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="20,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#444444"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                CornerRadius="10"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}">
                            <Grid>
                                <Rectangle Height="14"
                                           VerticalAlignment="Top"
                                           RadiusX="10"
                                           RadiusY="10"
                                           Opacity="0.22"
                                           Fill="White"/>
                                <ContentPresenter HorizontalAlignment="Center"
                                                  VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#3A3A49"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#8A2BE2"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Warning Icon and Header -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,20">
            <TextBlock Text="!" FontSize="32" FontWeight="Bold" Foreground="#8A2BE2" Margin="0,0,15,0"/>
            <StackPanel>
                <TextBlock Text="Confirm Deletion" 
                          FontSize="20" 
                          FontWeight="Bold" 
                          Foreground="White"/>
                <TextBlock Text="This action cannot be undone" 
                          FontSize="12" 
                          Foreground="#A6A6A6" 
                          Margin="0,5,0,0"/>
            </StackPanel>
        </StackPanel>
        
        <!-- Content -->
        <Border Grid.Row="1" 
                Background="#2D2D39" 
                CornerRadius="10" 
                Padding="15"
                BorderThickness="1"
                BorderBrush="#444444">
            <StackPanel>
                <TextBlock Text="You are about to delete:" 
                          FontSize="14" 
                          FontWeight="SemiBold"
                          Foreground="#A6A6A6"
                          Margin="0,0,0,10"/>
                <TextBlock Name="ItemTypeText" 
                          FontSize="14" 
                          Foreground="#7367F0"
                          FontWeight="Bold"
                          Margin="0,0,0,15"/>
                <ScrollViewer VerticalScrollBarVisibility="Auto" 
                             HorizontalScrollBarVisibility="Auto">
                    <TextBlock Name="VMNamesText" 
                              FontSize="13" 
                              TextWrapping="Wrap"
                              Foreground="#E5E7EB"
                              FontFamily="Consolas"
                              Padding="5"/>
                </ScrollViewer>
            </StackPanel>
        </Border>
        
        <!-- Buttons -->
        <StackPanel Grid.Row="2" 
                   Orientation="Horizontal" 
                   HorizontalAlignment="Right" 
                   Margin="0,20,0,0">
            <Button Name="CancelButton" 
                   Content="Cancel" 
                   Background="#3E3E4E" 
                   Foreground="#FFFFFF"
                   BorderBrush="#9A9AAF"/>
            <Button Name="DeleteButton" 
                   Content="Delete" 
                   Background="#EA5455" 
                   Foreground="#FFFFFF"
                   BorderBrush="#FF7A7A"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    
    # מציאת אלמנטים
    $itemTypeTextBlock = $window.FindName("ItemTypeText")
    $vmNamesTextBlock = $window.FindName("VMNamesText")
    $deleteButton = $window.FindName("DeleteButton")
    $cancelButton = $window.FindName("CancelButton")
    
    # הגדרת תוכן
    $itemTypeTextBlock.Text = $itemType
    $vmNamesTextBlock.Text = $vmNamesText
    
    # הגדרת פעולות לכפתורים
    $window.Add_Loaded({
        $deleteButton.Add_Click({
            $window.DialogResult = $true
            $window.Close()
        })
        
        $cancelButton.Add_Click({
            $window.DialogResult = $false
            $window.Close()
        })
    })
    
    # הצגת החלון
    $result = $window.ShowDialog()
    
    if (-not $result) {
        return
    }
}
    # Disable controls during deletion
    Set-ControlsEnabled -Enabled $false
    
    $totalVMs = $selectedVMs.Count
    $currentVM = 0
    
    # === רשימות לדוח סיכום ===
    $successList = @()
    $failedList = @()
    $skippedList = @()
    
    foreach ($vmInfo in $selectedVMs) {
        $vmName = $vmInfo.Name
        $osType = $vmInfo.OSType
        $vmEnvironment = $vmInfo.Environment
        $currentVM++
        $lblVMProgress.Text = "VM Progress: $vmName ($currentVM/$totalVMs)"
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        Write-Log "Starting deletion process for VM: $vmName ($currentVM of $totalVMs)"
        Reset-AllStatus

        # Track deletion results for this VM
        $vmResult = @{
            VMName = $vmName
            OSType = $osType
            Success = $true
            Errors = @()
            SCCMDeleted = $false
            ADDeleted = $false
            DNSDeleted = $false
            VMDeleted = $false
            Archived = $false
            EmailSent = $false
        }
        
        # Task 0: vCenter Connection (already connected)
        Update-TaskStatus -TaskIndex 0 -Status "Success" -Message "Connected to $($global:DefaultVIServer.Name)"
        $progressBar.Value = 14
        
        # Get VM details before deletion
        try {
            $vmDetails = Get-VMDetailsForArchive -VMName $vmName
        }
        catch {
            $vmResult.Errors += "Failed to get VM details: $($_.Exception.Message)"
            $vmDetails = @{ ServerName = $vmName; VMName = $vmName }
        }
        
        # Task 1: Delete SCCM Record
        if ($chkSCCM.Checked) {
            Update-TaskStatus -TaskIndex 1 -Status "Running" -Message ""
            try {
                $sccmResult = Remove-SCCMDevice -ComputerName $vmName -SCCMServer $script:DefaultSettings.SCCMServer -SiteCode $script:DefaultSettings.SCCMSiteCode
                
                if ($sccmResult.Success) {
                    $vmResult.SCCMDeleted = $sccmResult.Found
                    Update-TaskStatus -TaskIndex 1 -Status "Success" -Message $sccmResult.Message
                } else {
                    Update-TaskStatus -TaskIndex 1 -Status "Failed" -Message $sccmResult.Message
                    $vmResult.Errors += "SCCM: $($sccmResult.Message)"
                }
            }
            catch {
                Update-TaskStatus -TaskIndex 1 -Status "Failed" -Message $_.Exception.Message
                $vmResult.Errors += "SCCM: $($_.Exception.Message)"
            }
        } else {
            Update-TaskStatus -TaskIndex 1 -Status "Skipped" -Message "SCCM deletion skipped"
        }
        $progressBar.Value = 28
        
        # Task 2: Delete AD Record
        if ($chkAD.Checked) {
            Update-TaskStatus -TaskIndex 2 -Status "Running" -Message ""
            try {
                $adResult = Remove-ADDevice -ComputerName $vmName `
                    -DomainController $script:DefaultSettings.DomainController `
                    -SearchBase (Get-ADSearchBaseForEnvironment -EnvironmentName $vmEnvironment)
        
                if ($adResult.Success) {
                    $vmResult.ADDeleted = $adResult.Found
                    Update-TaskStatus -TaskIndex 2 -Status "Success" -Message $adResult.Message
                } else {
                    Update-TaskStatus -TaskIndex 2 -Status "Failed" -Message $adResult.Message
                    $vmResult.Errors += "AD: $($adResult.Message)"
                }
            }
            catch {
                Update-TaskStatus -TaskIndex 2 -Status "Failed" -Message $_.Exception.Message
                $vmResult.Errors += "AD: $($_.Exception.Message)"
            }
        } else {
            Update-TaskStatus -TaskIndex 2 -Status "Skipped" -Message "AD deletion skipped"
        }
        $progressBar.Value = 42
        
        # Task 3: Delete DNS Records
        if ($chkDNS.Checked) {
            Update-TaskStatus -TaskIndex 3 -Status "Running" -Message ""
            try {
                $dnsResult = Remove-DNSRecords -ComputerName $vmName -DNSServer $script:DefaultSettings.DNSServer -DNSZone $script:DefaultSettings.DNSZone
                
                if ($dnsResult.Success) {
                    $vmResult.DNSDeleted = $dnsResult.Found
                    Update-TaskStatus -TaskIndex 3 -Status "Success" -Message $dnsResult.Message
                } else {
                    Update-TaskStatus -TaskIndex 3 -Status "Failed" -Message $dnsResult.Message
                    $vmResult.Errors += "DNS: $($dnsResult.Message)"
                }
            }
            catch {
                Update-TaskStatus -TaskIndex 3 -Status "Failed" -Message $_.Exception.Message
                $vmResult.Errors += "DNS: $($_.Exception.Message)"
            }
        } else {
            Update-TaskStatus -TaskIndex 3 -Status "Skipped" -Message "DNS deletion skipped"
        }
        $progressBar.Value = 56
        
# Task 4: Delete VM or Template
if ($chkVM.Checked) {

    Update-TaskStatus -TaskIndex 4 -Status "Running" -Message ""

    try {
        # בדוק אם זה Template
        $vmData = $script:allVMsData | Where-Object { $_.Name -eq $vmName } | Select-Object -First 1
        $isTemplate = if ($vmData) { $vmData.IsTemplate } else { $false }
        
        if ($isTemplate) {
            # מחיקת Template
            Update-TaskStatus -TaskIndex 4 -Status "Running" -Message "Removing template..."
            
            $template = Get-Template -Name $vmName -ErrorAction Stop
            
            if ($script:DefaultSettings.DeleteFromDisk) {
                Remove-Template -Template $template -DeletePermanently -Confirm:$false -ErrorAction Stop
            }
            else {
                Remove-Template -Template $template -Confirm:$false -ErrorAction Stop
            }
            
            $vmResult.VMDeleted = $true
            Update-TaskStatus -TaskIndex 4 -Status "Success" -Message "Template deleted successfully"
        }
        else {
            # מחיקת VM רגיל
            $vm = Get-VM -Name $vmName -ErrorAction Stop

            # Power off if running
            if ($vm.PowerState -eq "PoweredOn") {

                Update-TaskStatus -TaskIndex 4 -Status "Running" -Message "Powering off VM..."

                try {
                    Stop-VM -VM $vm -Confirm:$false -ErrorAction Stop | Out-Null

                    $timeout = $script:DefaultSettings.PowerOffTimeout
                    $elapsed = 0

                    while ((Get-VM -Name $vmName -ErrorAction SilentlyContinue).PowerState -eq "PoweredOn" -and $elapsed -lt $timeout) {
                        Start-Sleep -Seconds 2
                        $elapsed += 2
                    }
                }
                catch {
                    $vmResult.Errors += "Power off failed: $($_.Exception.Message)"
                }

            }

            # Delete VM
            if ($script:DefaultSettings.DeleteFromDisk) {
                Remove-VM -VM $vm -DeletePermanently -Confirm:$false -ErrorAction Stop
            }
            else {
                Remove-VM -VM $vm -Confirm:$false -ErrorAction Stop
            }

            $vmResult.VMDeleted = $true
            Update-TaskStatus -TaskIndex 4 -Status "Success" -Message "VM deleted successfully"
        }
    }
    catch {

        $vmResult.VMDeleted = $false
        $vmResult.Success = $false
        $vmResult.Errors += "VM/Template Delete: $($_.Exception.Message)"

        Update-TaskStatus -TaskIndex 4 -Status "Failed" -Message $_.Exception.Message
        Write-Log "VM/Template deletion failed for $vmName : $($_.Exception.Message)" -Level "ERROR"

    }

}

        $progressBar.Value = 70

        # Task 5: Archive to SQL
        if ($chkArchive.Checked -and $script:sqlTableExists) {
            Update-TaskStatus -TaskIndex 5 -Status "Running" -Message ""
            try {
                $status = if ($vmResult.Success) { "Success" } else { "Failed" }

                $archiveResult = Save-VMDeletionToArchive -VMDetails $vmDetails -Reason $reason `
                    -SCCMDeleted $vmResult.SCCMDeleted -ADDeleted $vmResult.ADDeleted `
                    -DNSDeleted $vmResult.DNSDeleted -VMDeleted $vmResult.VMDeleted -Status $status

                if ($archiveResult.Success) {
                    $vmResult.Archived = $true
                    Update-TaskStatus -TaskIndex 5 -Status "Success" -Message $archiveResult.Message
                }
                else {
                    $vmResult.Errors += "Archive: $($archiveResult.Message)"
                    Update-TaskStatus -TaskIndex 5 -Status "Failed" -Message $archiveResult.Message
                }
            }
            catch {
                $vmResult.Errors += "Archive: $($_.Exception.Message)"
                Update-TaskStatus -TaskIndex 5 -Status "Failed" -Message $_.Exception.Message
                Write-Log "Archive failed for $vmName : $($_.Exception.Message)" -Level "ERROR"
            }
        }
        else {
            Update-TaskStatus -TaskIndex 5 -Status "Skipped" -Message "Archive skipped"
        }
        $progressBar.Value = 85

        # Task 6: Send Email Notification
        if ($chkEmail.Checked -and $script:DefaultSettings.EnableEmailNotification) {
            Update-TaskStatus -TaskIndex 6 -Status "Running" -Message ""
            try {
                $emailResult = Send-VMDeletionNotification -VMDetails $vmDetails -Reason $reason -OSType $osType -DeletionSuccess $vmResult.VMDeleted
                
                if ($emailResult.Success) {
                    $vmResult.EmailSent = $true
                    Update-TaskStatus -TaskIndex 6 -Status "Success" -Message $emailResult.Message
                } else {
                    $vmResult.Errors += "Email: $($emailResult.Message)"
                    Update-TaskStatus -TaskIndex 6 -Status "Failed" -Message $emailResult.Message
                }
            }
            catch {
                $vmResult.Errors += "Email: $($_.Exception.Message)"
                Update-TaskStatus -TaskIndex 6 -Status "Failed" -Message $_.Exception.Message
            }
        } else {
            Update-TaskStatus -TaskIndex 6 -Status "Skipped" -Message "Email notification skipped"
        }
        $progressBar.Value = 100
        
        # הוסף לרשימה המתאימה
        if ($vmResult.VMDeleted -or (-not $chkVM.Checked -and $vmResult.Errors.Count -eq 0)) {
            $successList += $vmResult
        }
        elseif ($vmResult.Errors.Count -gt 0) {
            $failedList += $vmResult
        }
        else {
            $skippedList += $vmResult
        }
        
        Write-Log "Deletion process completed for VM: $vmName (Success: $($vmResult.Success))"
        
        # המתן קצת בין מכונות
        Start-Sleep -Milliseconds 500
        }
    $lblVMProgress.Text = "VM Progress: Completed"
    $lblVMProgress.ForeColor = [System.Drawing.Color]::FromArgb(40, 167, 69)

    # === שלח דוח סיכום במייל ===
    if ($failedList.Count -gt 0 -or $successList.Count -gt 0) {
        Send-DeletionSummaryReport -SuccessList $successList -FailedList $failedList -SkippedList $skippedList -Reason $reason
    }

    # Re-enable controls
    Set-ControlsEnabled -Enabled $true
    
    # Refresh VM list
    Get-VMList
   
    # איפוס Progress Bar וסטטוסים
    $progressBar.Value = 0
    Reset-AllStatus
    
    # נקה את הבחירות
    $script:SelectedVMNames.Clear()
    $lblSelectedCount.Text = "0 selected"
    
    # בטל סימון כל הפריטים ברשימה
    $script:IsUpdatingList = $true
    foreach ($item in $lstVMs.Items) {
        $item.Checked = $false
    }
    $script:IsUpdatingList = $false
    
    # הודעת סיכום - TopMost ראשון במסך
    $summaryMsg = "Deletion process completed!`n`n"
    $summaryMsg += "Successful: $($successList.Count)`n"
    $summaryMsg += "Failed: $($failedList.Count)`n"
    $summaryMsg += "Skipped: $($skippedList.Count)`n"

    if ($failedList.Count -gt 0) {
        $summaryMsg += "`nFailed VMs:`n"
        foreach ($failed in $failedList) {
            $summaryMsg += "  - $($failed.VMName): $($failed.Errors -join '; ')`n"
        }
    }
    
    $icon = if ($failedList.Count -eq 0) { 
        [System.Windows.Forms.MessageBoxIcon]::Information 
    } else { 
        [System.Windows.Forms.MessageBoxIcon]::Warning 
    }
    
    # הצג הודעה TopMost בראש המסך
    $topForm = New-Object System.Windows.Forms.Form
    $topForm.TopMost = $true
    $topForm.StartPosition = "CenterScreen"
    $topForm.ShowInTaskbar = $false
    $topForm.Size = New-Object System.Drawing.Size(1, 1)
    $topForm.Opacity = 0
    $topForm.Show()
    
    [System.Windows.Forms.MessageBox]::Show(
        $topForm,
        $summaryMsg,
        "Deletion Summary",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    )
    
    $topForm.Close()
    $topForm.Dispose()
    
    # החזר פוקוס לחלון הראשי
    $form.Activate()
})

# ========== Resume Button ==========
$btnResume.Add_Click({
    $progressFile = "$env:TEMP\VMDeletionTool\progress.json"
    $availableProgressFile = Wait-JsonFileAvailable -Path $progressFile
    if ($availableProgressFile) {
        $progress = Get-Content -LiteralPath $availableProgressFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Resume from: $($progress.LastVM)?`nLast run: $($progress.Timestamp)",
            "Resume Deletion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            # מצא את ה-VM ברשימה וסמן אותו
            $script:SelectedVMNames.Clear()
            $script:SelectedVMNames.Add($progress.LastVM) | Out-Null
            
            # סמן ברשימה
            foreach ($item in $lstVMs.Items) {
                if ($item.Text -eq $progress.LastVM) {
                    $item.Selected = $true
                    $lstVMs.EnsureVisible($item.Index)
                    break
                }
            }
            
            $lblSelectedCount.Text = "1 selected"
            
            [System.Windows.Forms.MessageBox]::Show(
                "VM '$($progress.LastVM)' selected.`nClick 'Delete VM' to continue.",
                "Ready to Resume",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
    }
    else {
        [System.Windows.Forms.MessageBox]::Show(
            "No progress file found. Nothing to resume.",
            "Resume",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
})

# ========== Close Button ==========
$btnClose.Add_Click({
    if ($script:vcConnected) {
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Disconnect from vCenter and close?",
            "Confirm Close",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
            Disconnect-VIServer -Server * -Force -Confirm:$false -ErrorAction SilentlyContinue
            $form.Close()
        }
    } else {
        $form.Close()
    }
})

# ========== Form Closing Event ==========
$form.Add_FormClosing({})
#endregion Event Handlers

# SQL is checked manually from the Archive tab to keep GUI startup fast.
$form.Add_Shown({
    $lblArchiveStatus.Text = "SQL not checked yet. Click 'Check Table' when needed."
    $lblArchiveStatus.ForeColor = $uiMuted
})

# Show the form
[void]$form.ShowDialog()

