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

function Start-SC2Deploy {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [switch]$UpdateWindows
    )
    #=======================================================================
    #	Block
    #=======================================================================
    Block-StandardUser
    Block-WindowsVersionNe10
    Block-PowerShellVersionLt5
    #=======================================================================
    #   Header
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Green 'Start-SC2 Update Deploy'
    #=======================================================================
    #   Transcript
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Start-Transcript"
    $Transcript = "$((Get-Date).ToString('yyyy-MM-dd-HHmmss'))-SC2Deploy.log"
    Start-Transcript -Path (Join-Path "$env:SystemRoot\Temp" $Transcript) -ErrorAction Ignore
    Write-Host -ForegroundColor DarkGray '========================================================================='
    #=======================================================================
    #=======================================================================
    #   PSGallery
    #=======================================================================
    $PSGalleryIP = (Get-PSRepository -Name PSGallery).InstallationPolicy
    if ($PSGalleryIP -eq 'Untrusted') {
        Write-Host -ForegroundColor DarkGray '========================================================================='
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Set-PSRepository -Name PSGallery -InstallationPolicy Trusted"
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
    }
    #=======================================================================
    #=======================================================================
    #	Windows Update Software
    #=======================================================================
    if ($UpdateWindows) {
        Write-Host -ForegroundColor DarkGray '========================================================================='
        Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Windows and Microsoft Update Software"
        if (!(Get-Module PSWindowsUpdate -ListAvailable)) {
            try {
                Install-Module PSWindowsUpdate -Force
            } catch {
                Write-Warning 'Unable to install PSWindowsUpdate PowerShell Module'
                $UpdateWindows = $false
            }
        }
    }
    if ($UpdateWindows) {
        Write-Host -ForegroundColor DarkCyan 'Add-WUServiceManager -MicrosoftUpdate -Confirm:$false'
        Add-WUServiceManager -MicrosoftUpdate -Confirm:$false
        Write-Host -ForegroundColor DarkCyan 'Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -NotTitle Malicious -UpdateType Driver'
        Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot -NotTitle 'Malicious' -UpdateType 'Driver'
    }
    #=======================================================================
    #	Stop-Transcript
    #=======================================================================
    Write-Host -ForegroundColor DarkGray '========================================================================='
    Write-Host -ForegroundColor Cyan "$((Get-Date).ToString('yyyy-MM-dd-HHmmss')) Stop-Transcript"
    Stop-Transcript
    Write-Host -ForegroundColor DarkGray '========================================================================='
    #=======================================================================
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

    $InstalledOperatingSystemLabel = New-Object system.Windows.Forms.Label
    $InstalledOperatingSystemLabel.AutoSize = $true
    $InstalledOperatingSystemLabel.text = "$Windows $Edition $InstalledBuild"
    $InstalledOperatingSystemLabel.width = 780
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

    $SerialLabel = New-Object system.Windows.Forms.Label
    $SerialLabel.AutoSize = $true
    $SerialLabel.text = "$Serial"
    $SerialLabel.width = 780
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
    $UpdateWindowsButton.text = 'Update Drivers'
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
        Start-SC2Deploy -UpdateWindows -Verbose
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