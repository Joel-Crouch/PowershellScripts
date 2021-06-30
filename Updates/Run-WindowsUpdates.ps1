<#
.SYNOPSIS
    Install windows updates via powershell

.DESCRIPTION
	Installs PSWindows update module, checks for updates and installs. Used for local installs

.PARAMETER Restart
	Automatically reboot after updates are installed.

.PARAMETER Drivers
	Checks and updates drivers.
#>
[CmdletBinding()]
param (
	[switch]$Restart,
    [switch]$Drivers
)

#Functions
function Load-PSWindowsUpdate {
    $ModuleError = $null
	Try {
		import-module PSWindowsUpdate -ErrorAction Stop
	}catch{
        $ModuleError = $true
		Write-Warning "PSWindowsUpdate not found. Installing module"
					
		if (([System.Environment]::OSVersion.Version).Major -ne 10){
			Write-Error "Windows 10 required"
		}else{
			Install-PackageProvider NuGet -Force | Out-Null
			Set-PSRepository PSGallery -InstallationPolicy Trusted | Out-Null
			Install-Module PSWindowsUpdate -force -confirm:$false | Out-Null
		}
	}
    if ($ModuleError -ne $true){
        #check for updates
        $CurrentAWSPSModule = ((Get-Module -Name PSWindowsUpdate -ListAvailable).Version | Sort-Object -Descending | Select-Object -First 1).ToString()
        $NewestAWSPSModule = (Find-Module -Name PSWindowsUpdate).Version.ToString()
        if ([System.Version]$CurrentAWSPSModule -lt [System.Version]$NewestAWSPSModule){
            Write-Verbose "Module is out of date. Attempting to update" -verbose
            Update-Module PSWindowsUpdate -force -confirm:$false | Out-Null
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

#Run Updates
if ($Restart){
	Install-WindowsUpdate -NotCategory $UpdateCats -NotTitle 'Feature|Preview' -MicrosoftUpdate -AcceptAll -AutoReboot | Out-File $env:ALLUSERSPROFILE\AdminScripts\PSWindowsUpdate.log
}else{
	Install-WindowsUpdate -NotCategory $UpdateCats -NotTitle 'Feature|Preview' -MicrosoftUpdate -AcceptAll -IgnoreReboot | Out-File $env:ALLUSERSPROFILE\AdminScripts\PSWindowsUpdate.log
}
