<#
.SYNOPSIS
    Simple script for reseting student passwords on Active Directory

.DESCRIPTION
    This script can be used to deploy to staff to help delegate student passwords.
Script is set to loop for multiple runs. Password format might need to be adjusted based on organization starting around line 62.

.PARAMETER StudentOU
    Root OU path that student accounts are located. Format needs to be distinguished name (OU=Students,DC=domain,DC=net)

.PARAMETER SkywardData
    Path to csv export of students from skyward. This export is used to look up lunch pins for students.

.PARAMETER SkywardPWDefault
    String used in conjunction with import of skyward data (e.g. sample12345).

.PARAMETER ResetLog
    Path to save logs of students and staff that were used in script. Accounts running the script need write access to path.

.EXAMPLE
    .'.\Reset-StudentPassword' -StudentOU 'OU=Students,DC=domain,DC=net' -SkywardData '\\server.domain.org\skyward\students.csv' -SkywardPWDefault 'apple' -ResetLog '\\server.domain.org\logs\StudentReset.log'
#>
[CmdletBinding()]
param (
    [string]$StudentOU,
    [string]$SkywardData,
    [string]$SkywardPWDefault,
    [string]$ResetLog
)

#Require parameter
if (!$StudentOU){
    do {
        $StudentOU = Read-Host "Supply values for the following parameters:`nStudentOU"
    }until ($StudentOU)
}

#Pull skyward data
if (($SkywardData) -and (Test-Path $SkywardData)){
    $SkywardStudents = Import-Csv -Path $SkywardData
}

While ($true){
    $StudentInput = (Read-Host 'Username of student').Trim()

    #ignore if email address was put in
    if ($StudentInput -like '*@*'){
        $StudentInput = ($StudentInput -split '@')[0]
        Write-Host "Only username is required. Proceeding with only $StudentInput for username." -ForegroundColor Cyan
    }

    #Verify student exists
    $Searcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
    $Searcher.SearchRoot = [ADSI]"LDAP://$StudentOU"
    $Searcher.Filter = "(&(objectCategory=user)(|(sAMAccountName=$StudentInput)(cn=$StudentInput)))"
    $LDAPObject = $Searcher.FindOne()
    if ($LDAPObject){
        do{
            $Fault = $null
            do{
                ################################# Adjust based on organization password defaults #################################
                #find student default password
                if (($SkywardData) -and (Test-Path $SkywardData)){
                    $LunchPin=($SkywardStudents | Where-Object {($_.'Schl Email Addr' -split '@')[0] -eq $StudentInput}).'Lunch Pin Num'
                    $DefaultPassword = $SkywardPWDefault+$LunchPin
                }
                if ($LunchPin){
                    $PasswordInput = Read-Host "Press enter to reset to default password of $DefaultPassword or type new password"
                    if (!$PasswordInput){
                        $PasswordInput = $DefaultPassword
                    }
                }else{
                    $PasswordInput = Read-Host 'New student password'
                }
                if ($PasswordInput.Length -lt 8){
                    Write-Host 'Minimum password length is 8 characters' -ForegroundColor Red
                }
            }until($PasswordInput.Length -gt 7)
            try{
                ([ADSI]$LDAPObject.Path).SetPassword($PasswordInput)
            }catch{
                $Fault = $true
                if ($Error[0] -match 'The server is unwilling to process the request'){
                    Write-Host 'The password was not accepted by Active Directory.`nTry a different password.' -ForegroundColor Red
                }else{
                    Write-Host $Error[0] -ForegroundColor Red
                }
            
            }
        }until ($null -eq $Fault)
        Write-Host "Password was set successfully for $StudentInput" -ForegroundColor Green
        "$(Get-Date) - $env:USERNAME on $env:COMPUTERNAME reset password for $StudentInput" | Out-File -Append -FilePath $ResetLog
    }else{
        Write-Host -ForegroundColor Red "No student account was found for $StudentInput. Verify correct username and contact IS helpdesk if it is correct."
    }
}
