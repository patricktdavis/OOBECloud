##=============================================================================
#region SCRIPT DETAILS
#=============================================================================

<#
.SYNOPSIS
Runs a menu on OOBE Screen allowing techs to run Windows update, Update drivers and add Windows capabilities
.EXAMPLE
PS C:\> OOBECloud-Hennepin.ps1
#>

#=============================================================================
#endregion
#=============================================================================

$Serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber
If (Test-Path -Path 'C:\Windows\System32\kernel32.dll') {
    $Edition = (Get-WindowsEdition -Online).edition
    $BuildNumber = ([environment]::OSVersion.Version).Build
    $Windows = 'Windows 10'

    if ($BuildNumber -eq '19042') {
        $InstalledBuild = '20H2'
    } elseif ($BuildNumber -eq '19041') {
        $InstalledBuild = '20H2'
    } elseif ($BuildNumber -eq '18363') {
        $InstalledBuild = '1909'
    } elseif ($BuildNumber -eq '18362') {
        $InstalledBuild = '1903'
    } elseif ($BuildNumber -eq '17763') {
        $InstalledBuild = '1809'
    } elseif ($BuildNumber -eq '17134') {
        $InstalledBuild = '1803'
    } elseif ($BuildNumber -eq '16299') {
        $InstalledBuild = '1709'
    }
} else {
    $InstalledBuild = 'Installed'
    $Edition = 'No OS'
    $Windows = ''
}

#=============================================================================
#region FUNCTIONS
#=============================================================================
function Invoke-OOBECloud {
    <#
    .SYNOPSIS
    Synopsis
    .EXAMPLE
    Invoke-OOBECloud
    .INPUTS
    None
    You cannot pipe objects to Invoke-OOBECloud.
    .OUTPUTS
    None
    The cmdlet does not return any output.
    #>

    [CmdletBinding()]
    Param()

    if ((Get-MyComputerModel) -match 'Virtual') {
        Write-Host -ForegroundColor Cyan 'Setting Display Resolution to 1600x'
        Set-DisRes 1600
    }
    Invoke-NewOOBEBoxHD
}


function Invoke-NewOOBEBoxHD {
    <#
    .SYNOPSIS
    Synopsis

    .EXAMPLE
    Invoke-NewOOBEBoxHD

    .INPUTS
    None
    You cannot pipe objects to Invoke-NewOOBEBoxHD.

    .OUTPUTS
    None
    The cmdlet does not return any output.
    #>

    [CmdletBinding()]
    Param()

    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()

    $OOBECloud = New-Object system.Windows.Forms.Form
    $OOBECloud.ClientSize = New-Object System.Drawing.Point(800,450)
    $OOBECloud.text = 'Hennepin County SSD Team'
    $OOBECloud.TopMost = $true

    $Title = New-Object system.Windows.Forms.Label
    $Title.text = 'Welcome to Hennepin County Imaging'
    $Title.AutoSize = $true
    $Title.width = 25
    $Title.height = 10
    $Title.location = New-Object System.Drawing.Point(26,34)
    $Title.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $InstalledOSDescription = New-Object system.Windows.Forms.Label
    $InstalledOSDescription.text = 'The following Operating System is installed: '
    $InstalledOSDescription.AutoSize = $true
    $InstalledOSDescription.width = 25
    $InstalledOSDescription.height = 10
    $InstalledOSDescription.location = New-Object System.Drawing.Point(43,119)
    $InstalledOSDescription.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

    $InstalledOperatingSystemLabel = New-Object system.Windows.Forms.TextBox
    $InstalledOperatingSystemLabel.multiline = $false
    $InstalledOperatingSystemLabel.text = "$Windows $Edition $InstalledBuild"
    $InstalledOperatingSystemLabel.width = 100
    $InstalledOperatingSystemLabel.height = 20
    $InstalledOperatingSystemLabel.location = New-Object System.Drawing.Point(41,163)
    $InstalledOperatingSystemLabel.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $SerialDescription = New-Object system.Windows.Forms.Label
    $SerialDescription.text = 'This Machines Serial Number is:'
    $SerialDescription.AutoSize = $true
    $SerialDescription.width = 25
    $SerialDescription.height = 10
    $SerialDescription.location = New-Object System.Drawing.Point(43,220)
    $SerialDescription.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

    $SerialLabel = New-Object system.Windows.Forms.TextBox
    $SerialLabel.multiline = $false
    $SerialLabel.text = "$Serial"
    $SerialLabel.width = 100
    $SerialLabel.height = 20
    $SerialLabel.location = New-Object System.Drawing.Point(43,260)
    $SerialLabel.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $MakeASelection = New-Object system.Windows.Forms.Label
    $MakeASelection.text = 'Please make a selection below:'
    $MakeASelection.AutoSize = $true
    $MakeASelection.width = 25
    $MakeASelection.height = 10
    $MakeASelection.location = New-Object System.Drawing.Point(39,300)
    $MakeASelection.Font = New-Object System.Drawing.Font('Segoe UI',20)

    $ExitButton = New-Object system.Windows.Forms.Button
    $ExitButton.text = 'Exit'
    $ExitButton.width = 200
    $ExitButton.height = 43
    $ExitButton.location = New-Object System.Drawing.Point(43,360)
    $ExitButton.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $ExitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $UpdateWindowsButton = New-Object system.Windows.Forms.Button
    $UpdateWindowsButton.text = 'Update Windows 10'
    $UpdateWindowsButton.width = 400
    $UpdateWindowsButton.height = 43
    $UpdateWindowsButton.location = New-Object System.Drawing.Point(300,360)
    $UpdateWindowsButton.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $UpdateWindowsButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $OOBECloud.controls.AddRange(@($Title,$InstalledOSDescription,$InstalledOperatingSystemLabel,$SerialDescription,$SerialLabel,$MakeASelection,$ExitButton,$UpdateWindowsButton))


    $Result = $OOBECloud.ShowDialog()

    If ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        #Installing latest OSD Content
        Write-Host -ForegroundColor Cyan 'Updating OSD PowerShell Module'
        Install-Module OSD -Force
        Write-Host -ForegroundColor Cyan 'Importing OSD PowerShell Module'
        Import-Module OSD -Force
        Start-OOBEDeploy -AddNetFX3 -UpdateDrivers -UpdateWindows -Verbose
    }
}

#=============================================================================
#endregion
#=============================================================================
#region EXECUTION
#=============================================================================

#Running Popup
Invoke-OOBECloud
#Reset Execution Policy
Write-Host -ForegroundColor Cyan 'Resetting PS Execution Policy'
Set-ExecutionPolicy Restricted -Force -Verbose
#Restart Machine
Write-Host -ForegroundColor Cyan 'Restarting in 20 seconds!'
Start-Sleep -Seconds 20
Restart-Computer

#=============================================================================
#endregion
#=============================================================================
