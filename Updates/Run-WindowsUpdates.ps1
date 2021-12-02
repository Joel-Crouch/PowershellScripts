<#
.SYNOPSIS
    Install windows updates via powershell

.DESCRIPTION
	Installs PSWindows update module, checks for updates and installs. Used for local installs

.PARAMETER Restart
	Automatically reboot after updates are installed.

.PARAMETER Drivers
	Checks for drivers long with windows updates.

.PARAMETER DriversOnly
	Only checks for drivers
#>
[CmdletBinding()]
param (
	[switch]$Restart,
    [switch]$Drivers,
    [switch]$DriversOnly
)

#Functions
function Load-PSWindowsUpdate {
    #Force TLS Verion
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $ModuleError = $null
	$NuGet = Get-PackageProvider | where {$_.Name -like 'NuGet'}
    if (!$NuGet){
        Write-Warning "Installing NuGet Package Provider"
        Install-PackageProvider NuGet -Force | Out-Null
		Set-PSRepository PSGallery -InstallationPolicy Trusted | Out-Null
    }
	Try {
		import-module PSWindowsUpdate -ErrorAction Stop
	}catch{
        $ModuleError = $true
		Write-Warning "PSWindowsUpdate not found. Installing module"					
		Install-Module PSWindowsUpdate -force -confirm:$false | Out-Null

	}
    if ($ModuleError -ne $true){
        #check for updates
        Write-Verbose "Checking for module updates" -Verbose
        $CurrentAWSPSModule = ((Get-Module -Name PSWindowsUpdate -ListAvailable).Version | Sort-Object -Descending | Select-Object -First 1).ToString()
        $NewestAWSPSModule = (Find-Module -Name PSWindowsUpdate).Version.ToString()
        if ([System.Version]$CurrentAWSPSModule -lt [System.Version]$NewestAWSPSModule){
            Write-Verbose "Module is out of date. Attempting to update" -verbose
            Update-Module PSWindowsUpdate -force -confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
        
    }
}


#Variables
if (!(Test-Path $env:ALLUSERSPROFILE\AdminScripts)){
	New-Item $env:ALLUSERSPROFILE\AdminScripts -ItemType Directory | Out-Null
}else{
	Remove-Item $env:ALLUSERSPROFILE\AdminScripts\PSWindowsUpdate.log -ErrorAction SilentlyContinue
}

if ($Drivers){
	$UpdateCats = 'Feature Packs'
}else{
	$UpdateCats = 'Drivers','Feature Packs'
}

#Check status of module
Load-PSWindowsUpdate

$InstallParams = @{
    MicrosoftUpdate = $true
    AcceptAll = $true
    NotCategory = @('Feature Packs')
    NotTitle = 'Feature|Preview'
}

if ($Restart){
    $InstallParams.AutoReboot = $true
}else{
    $InstallParams.IgnoreReboot = $true
}

if ($DriversOnly){
    $InstallParams.Category = 'Drivers'
}elseif ($Drivers -eq $false){
    $InstallParams.NotCategory += 'Drivers'
}



#Run Updates
Install-WindowsUpdate @InstallParams | Out-File $env:ALLUSERSPROFILE\AdminScripts\PSWindowsUpdate.log
