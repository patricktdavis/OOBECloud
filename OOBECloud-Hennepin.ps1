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
    $Edition = (Get-WindowsEdition -Path c:\).edition
    Function Invoke-OSVersion {

        $signature = @'
[DllImport("kernel32.dll")]
public static extern uint GetVersion();
'@
        Add-Type -MemberDefinition $signature -Name 'Win32OSVersion' -Namespace Win32Functions -PassThru
    }
    $OSBuild = [System.BitConverter]::GetBytes((Invoke-OSVersion)::GetVersion())
    $Build = [byte]$OSBuild[2],[byte]$OSBuild[3]
    $BuildNumber = [System.BitConverter]::ToInt16($build,0)

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

    Add-Type -AssemblyName System.Windows.Forms
    $Monitors = [System.Windows.Forms.Screen]::AllScreens
    foreach ($Monitor in $Monitors) {
        $Width = $Monitor.bounds.Width
    }

    If ($Width -eq '3840') {
        Invoke-NewOOBEBox4K
    }

    Else {
        Invoke-NewOOBEBoxHD
    }
}

function Invoke-NewOOBEBox4K {
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
    $OOBECloud.ClientSize = New-Object System.Drawing.Point(737,400)
    $OOBECloud.text = 'Hennepin County SSD Team'
    $OOBECloud.TopMost = $true

    $Title = New-Object system.Windows.Forms.TextBox
    $Title.multiline = $false
    $Title.text = 'Welcome to Hennepin County Imaging'
    $Title.width = 100
    $Title.height = 20
    $Title.location = New-Object System.Drawing.Point(18,22)
    $Title.Font = New-Object System.Drawing.Font('Segoe UI',28,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $OSDescription = New-Object system.Windows.Forms.TextBox
    $OSDescription.multiline = $false
    $OSDescription.text = 'The following Operating System is installed: '
    $OSDescription.width = 100
    $OSDescription.height = 20
    $OSDescription.location = New-Object System.Drawing.Point(22,75)
    $OSDescription.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

    $OperatingSystemLabel = New-Object system.Windows.Forms.TextBox
    $OperatingSystemLabel.multiline = $false
    $OperatingSystemLabel.text = "Windows 10 $Edition $InstalledBuild"
    $OperatingSystemLabel.width = 100
    $OperatingSystemLabel.height = 20
    $OperatingSystemLabel.location = New-Object System.Drawing.Point(21,108)
    $OperatingSystemLabel.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $SerialDescription = New-Object system.Windows.Forms.TextBox
    $SerialDescription.multiline = $false
    $SerialDescription.text = 'This Machines Serial Number is:'
    $SerialDescription.width = 100
    $SerialDescription.height = 20
    $SerialDescription.location = New-Object System.Drawing.Point(21,162)
    $SerialDescription.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

    $SerialLabel = New-Object system.Windows.Forms.TextBox
    $SerialLabel.multiline = $false
    $SerialLabel.text = "$Serial"
    $SerialLabel.width = 100
    $SerialLabel.height = 20
    $SerialLabel.location = New-Object System.Drawing.Point(22,194)
    $SerialLabel.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $MakeASelection = New-Object system.Windows.Forms.TextBox
    $MakeASelection.multiline = $false
    $MakeASelection.text = 'Please make a selection below:'
    $MakeASelection.width = 100
    $MakeASelection.height = 20
    $MakeASelection.location = New-Object System.Drawing.Point(18,257)
    $MakeASelection.Font = New-Object System.Drawing.Font('Segoe UI',20)

    $ExitButton = New-Object system.Windows.Forms.Button
    $ExitButton.text = 'Exit'
    $ExitButton.width = 257
    $ExitButton.height = 43
    $ExitButton.location = New-Object System.Drawing.Point(21,316)
    $ExitButton.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $ExitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $UpdateWindowsButton = New-Object system.Windows.Forms.Button
    $UpdateWindowsButton.text = 'Update Windows 10'
    $UpdateWindowsButton.width = 274
    $UpdateWindowsButton.height = 43
    $UpdateWindowsButton.location = New-Object System.Drawing.Point(313,316)
    $UpdateWindowsButton.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $UpdateWindowsButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $OOBECloud.controls.AddRange(@($Title,$OSDescription,$OperatingSystemLabel,$SerialDescription,$SerialLabel,$MakeASelection,$ExitButton,$UpdateWindowsButton))


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
    $OOBECloud.ClientSize = New-Object System.Drawing.Point(737,400)
    $OOBECloud.text = 'Hennepin County SSD Team'
    $OOBECloud.TopMost = $true

    $Title = New-Object system.Windows.Forms.TextBox
    $Title.multiline = $false
    $Title.text = 'Welcome to Hennepin County Imaging'
    $Title.width = 100
    $Title.height = 20
    $Title.location = New-Object System.Drawing.Point(18,22)
    $Title.Font = New-Object System.Drawing.Font('Segoe UI',28,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $OSDescription = New-Object system.Windows.Forms.TextBox
    $OSDescription.multiline = $false
    $OSDescription.text = 'The following Operating System is installed: '
    $OSDescription.width = 100
    $OSDescription.height = 20
    $OSDescription.location = New-Object System.Drawing.Point(22,75)
    $OSDescription.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

    $OperatingSystemLabel = New-Object system.Windows.Forms.TextBox
    $OperatingSystemLabel.multiline = $false
    $OperatingSystemLabel.text = "Windows 10 $Edition $InstalledBuild"
    $OperatingSystemLabel.width = 100
    $OperatingSystemLabel.height = 20
    $OperatingSystemLabel.location = New-Object System.Drawing.Point(21,108)
    $OperatingSystemLabel.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $SerialDescription = New-Object system.Windows.Forms.TextBox
    $SerialDescription.multiline = $false
    $SerialDescription.text = 'This Machines Serial Number is:'
    $SerialDescription.width = 100
    $SerialDescription.height = 20
    $SerialDescription.location = New-Object System.Drawing.Point(21,162)
    $SerialDescription.Font = New-Object System.Drawing.Font('Segoe UI',14,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Italic))

    $SerialLabel = New-Object system.Windows.Forms.TextBox
    $SerialLabel.multiline = $false
    $SerialLabel.text = "$Serial"
    $SerialLabel.width = 100
    $SerialLabel.height = 20
    $SerialLabel.location = New-Object System.Drawing.Point(22,194)
    $SerialLabel.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))

    $MakeASelection = New-Object system.Windows.Forms.TextBox
    $MakeASelection.multiline = $false
    $MakeASelection.text = 'Please make a selection below:'
    $MakeASelection.width = 100
    $MakeASelection.height = 20
    $MakeASelection.location = New-Object System.Drawing.Point(18,257)
    $MakeASelection.Font = New-Object System.Drawing.Font('Segoe UI',20)

    $ExitButton = New-Object system.Windows.Forms.Button
    $ExitButton.text = 'Exit'
    $ExitButton.width = 257
    $ExitButton.height = 43
    $ExitButton.location = New-Object System.Drawing.Point(21,316)
    $ExitButton.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $ExitButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $UpdateWindowsButton = New-Object system.Windows.Forms.Button
    $UpdateWindowsButton.text = 'Update Windows 10'
    $UpdateWindowsButton.width = 274
    $UpdateWindowsButton.height = 43
    $UpdateWindowsButton.location = New-Object System.Drawing.Point(313,316)
    $UpdateWindowsButton.Font = New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]([System.Drawing.FontStyle]::Bold))
    $UpdateWindowsButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $OOBECloud.controls.AddRange(@($Title,$OSDescription,$OperatingSystemLabel,$SerialDescription,$SerialLabel,$MakeASelection,$ExitButton,$UpdateWindowsButton))


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
#Set Execution Policy
Write-Host -ForegroundColor Cyan 'Setting PS Execution Policy'
Set-ExecutionPolicy RemoteSigned -Force -Verbose
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
