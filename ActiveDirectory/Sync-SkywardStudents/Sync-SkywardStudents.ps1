<#
.SYNOPSIS
    Sync AD with Skyward

.DESCRIPTION
	Imports students to AD via an exported CSV

.PARAMETER Progress
	Shows progress bar.

.PARAMETER Domain
    Domain name

.PARAMETER DCServer
    DC server used to send AD requests. Required so that AD commands don't switch servers if there are multiple DCs.

.PARAMETER GCDSServer
    Server that Google Cloud Directory Sync is installed and configured

.PARAMETER WinSCPPortable
    Location of WinSCP executable used to download CSV from SFTP server

.PARAMETER TempDownload
    Location to download a copy of the CSV.

.PARAMETER StudentRootOU
    OU path for root student OU

.PARAMETER StudentGroupDN
    Distinguished name for the group used to keep track of accounts created by script.

.PARAMETER DoNotTrackGroupDN
    Distinguished name for the group used to exclude accounts from being tracked by script.

.PARAMETER Log
    Path in which to store script log

.PARAMETER ErrorEmailContact
    An email or list of emails to contact when an error occurs.

.PARAMETER SanityCount
    Minimum students. Used as a saftey guard incase of bad data pulled from export.

.PARAMETER SkipDownload
    Skip download from SFTP server

.PARAMETER NoEmail
    Do not send email for new accounts

.EXAMPLE
    PS> .Sync-SkywardStudents.ps1 -Progress:$false

#>
[CmdletBinding()]
param (
    [switch]$Progress = $true,
    [string]$Domain = "domain.net",
    [string]$DCServer = "dc.$Domain",
    [string]$GCDSServer = "sync-srv",
    [string]$WinSCPPortable = "\\server\software\WinSCP\WinSCP-5.17.7-Portable\WinSCP.com",
    [string]$TempDownload = "\\domain.net\dfs\AdminScripts\AD\Sync-SkywardStudents\",
    [string]$StudentRootOU = "OU=Students,OU=District Users,DC=domain,DC=net",
    [string]$StudentGroupDN = 'CN=GRP_StudentAccount,OU=Students,OU=District Users,DC=domain,DC=net',
    [string]$DoNotTrackGroupDN = 'CN=GRP_DoNotTrack,OU=District Users,DC=ohsd,DC=net',
    [string]$Log =  "\\domain.net\dfs\AdminScripts\AD\Sync-SkywardStudents\SkywardSyncLog.txt",
	$ErrorEmailContact = 'netops@domain.net',
    [int]$SanityCount = 5000,
    [switch]$SkipDownload,
    [switch]$NoEmail,
    [int]$WaitTime = 120
)

#Variables
$NewStudents = New-Object System.Collections.ArrayList
$DisabledStudents = New-Object System.Collections.ArrayList
$GraceStudents = New-Object System.Collections.ArrayList
$Today = Get-Date
$CurrentYear = Get-Date -Format 'yyyy'
$CurrentMonth = Get-Date -Format 'MM'
$GraceDate = (Get-Date).AddDays(14) | Get-Date -Format 'MM-dd-yy'
$ErrorEmail = $null

#Functions
Function Write-ErrorEmail {
    param (
        $ErrorString
    )
    $MailBody = @"
        <style>
            body,p,h3 { font-family: calibri; }
            h3  { margin-bottom: 5px; }
            th  { text-align: center; background: #003829; color: #FFF; padding: 5px; }
            td  { padding: 5px 20px; }
            tr  { background: #E7FFF9; }
        </style>

        <h3>Skyward Sync Errors:</h3>

        $ErrorString
        
"@
    Send-MailMessage -SmtpServer 'smtp-relay.gmail.com' -To $ErrorEmailContact -From 'noreply@domain.net' -Subject 'Account Creation Error' -Body $MailBody -BodyAsHtml
}

#Tables used to determine sub OUs basd on skyward entry Entity Name.
#MODIFY TO FIT YOUR OU STRUCTURE
#These OUs will all be children of $StudentRootOU
$StudentOUTable = @{
    'BV ELEMENTARY'                =  'OU=BVE,OU=Elementary,'
    'C HARBOR ELEMENTARY'          =  'OU=CHE,OU=Elementary,'
    'IGRAD ACADEMY'                =  'OU=iGrad Academy,OU=OHHS,'
    'HILLCREST ELEMENTARY'         =  'OU=HCE,OU=Elementary,'
    'HOMECONNECTION'               =  'OU=HC,'
    'NW MIDDLE SCHOOL'             =  'OU=NWMS,'
    'OH ELEMENTARY'                =  'OU=OHE,OU=Elementary,'
    'OH HIGH SCHOOL'               =  'OU=OHHS,'
    'OH INTERMEDIATE'              =  'OU=OHI,'
    'OH VIRTUAL ACADEMY'           =  'OU=OHVA,'
    'O VIEW ELEMENTARY'            =  'OU=OVE,OU=Elementary,'
    'Out of District'              =  'OU=Out of District,OU=Inactive Students,'
    'PRESCHOOL SPECIAL EDUCATION'  =  'OU=Preschool-Sped,'
    'RUNNING START'                =  'OU=Running Start,OU=OHHS,'
}

#Schools that do not have grad years in skyward
$NoGradYears = @(
    'IGRAD ACADEMY',
    'PRESCHOOL SPECIAL EDUCATION',
    'Out of District'
)
#sets up log
if(!(Test-Path $Log)){
        New-Item $Log -type file -Force > $null
}
#log start time
$StartTime = Get-Date
$Message = "`n$("*"*20) Script Started by $env:USERNAME on $env:COMPUTERNAME at $StartTime $("*"*20)"
Write-Host $Message -ForegroundColor Cyan
$Message | Out-File -FilePath $Log -Append

#Check if AD module loaded
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}catch{
    $ErrorMessage = 'AD module is required for this script. Please install and try again.'
    Write-Host $ErrorMessage -ForegroundColor Red
    Write-Output $ErrorMessage | Out-File $Log -Append
    if (!$NoEmail){
        Write-ErrorEmail -ErrorString "<p>$ErrorMessage</p>"
    }
    Exit 1
}

#Verify AD server is up
if (!(Test-Connection $DCServer -Quiet -Count 3)){
    $ErrorMessage = "Unable to connect to $DCServer."
    Write-Host $ErrorMessage -ForegroundColor Red
    Write-Output $ErrorMessage | Out-File $Log -Append
    if (!$NoEmail){
        Write-ErrorEmail -ErrorString "<p>$ErrorMessage</p>"
    }
    Exit 2
}

if (!$SkipDownload){
    if (!(test-path $TempDownload)){
        New-Item $TempDownload -ItemType Directory > $null
    }
    Write-Verbose "Downloading CSV" -Verbose

    #download student csv via SSH/SCP
    #SSH code can be generated by WinSCP
    #For security purposes, set account to read only permission
    $ExeOutput = & $WinSCPPortable `
      /ini=nul `
      /command `
        "open sftp://skyward:password@ftp.domain.net/ -hostkey=`"`"ssh-ed25519 255 3mxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx=`"`"" `
        "get Student.csv $TempDownload -resumesupport=off" `
        "exit"

    $winscpResult = $LASTEXITCODE
    if ($winscpResult -ne 0){
        $ErrorMessage = 'An error occurred while downloading Student CSV'
        Write-Host $ErrorMessage -ForegroundColor Red
        Write-Output $ExeOutput
        Write-Output $ErrorMessage,$ExeOutput | Out-File $Log -Append
        if (!$NoEmail){
            Write-ErrorEmail -ErrorString "<p>$ErrorMessage</p><pre>$ExeOutput</pre>"
        }
        Exit 3
    }
}

Write-Verbose 'Pulling Students from AD' -verbose
$SkywardStudents = Import-Csv -Path "$TempDownload\Student.csv"

#memberof is path to group used for tracking
$ADStudents = Get-AdUser -LDAPFilter "(&(objectCategory=user)(memberof=$StudentGroupDN))" `
                         -Properties Name,EmailAddress,DistinguishedName,Description,EmployeeID,Enabled,MemberOf,ObjectGUID `
                         -Server $DCServer -ErrorAction Stop

#Verify valid results
if ($SkywardStudents.Count -lt $SanityCount){
    if ($SkipDownload){
        Write-Warning "Student count is under $SanityCount."
        $LowStudentQ = Read-Host "Student count is $($SkywardStudents.Count). Are you sure you want to continue? y|n"
        if ($LowStudentQ -eq 'n'){
            Exit
        }
    }else{
        $ErrorMessage = "Student count is under $SanityCount. Skyward import shows $($SkywardStudents.Count). There may be a problem with the skyward CSV. Adjust or use parameter: SanityCount"
        Write-Host $ErrorMessage -ForegroundColor Red
        Write-Output $ErrorMessage | Out-File $Log -Append
        if (!$NoEmail){
            Write-ErrorEmail -ErrorString "<p>$ErrorMessage</p>"
        }
        Exit 4
    }
}

########################   BEGIN VERIFICATION OF STUDENT DATA   ########################
$i = 1
$TotalStudents = $SkywardStudents.Count
$Increments = [int]($TotalStudents / 100)
foreach ($Student in $SkywardStudents){
    if ($Progress -and ($i % $Increments -eq 0)){
        #Progress bar
        Write-Progress -Activity 'Processing Skyward Data' -Status "Student $i of $TotalStudents"  -PercentComplete ($i / $TotalStudents * 100)
    }
    
    #Variables
    $Email = $Student.'Schl Email Addr'.ToLower()
    $PWD = "default$($Student."Lunch Pin Num")"

    if (!$Email){
        #skip if no email address is found
        $ErrorMessage = "No email address was found for $($Student.'Stu Alphakey'). Check skyward for this student and assign an email address."
        Write-Warning $ErrorMessage
        Write-Output $ErrorMessage | Out-File $Log -Append
        $ErrorEmail = $ErrorEmail + "<p>$ErrorMessage</p>"
        continue
    }
    $Username = ($Email -split '@' )[0]

    #max length for samaccountname is 20
    if ($Username.length -gt 20){
        $SamAccount = $Username[0..19] -join ''
    }else{
        $SamAccount = $Username
    }

    #Verify entity exists in hash table
    if ($StudentOUTable[$Student.'Entity Name']){
        if ($NoGradYears -contains $Student.'Entity Name'){
            $ADOU = $StudentOUTable[$Student.'Entity Name'] + $StudentRootOU
            $NoGradYearStu = $True
        }else{
            $ADOU = "OU=$($Student.'Stu Grad Yr')," + $StudentOUTable[$Student.'Entity Name'] + $StudentRootOU
            #Check if OU exists for OUs with grad years
            if (! [adsi]::Exists("LDAP://$DCServer/$ADOU")){
                #OU does not exist create
                New-ADOrganizationalUnit -Name $($Student.'Stu Grad Yr') -Path ($StudentOUTable[$Student.'Entity Name'] + $StudentRootOU) `
                    -Server $DCServer
            }
            $NoGradYearStu = $false
        }

        #Default studentGroups
        $ADGroups = New-Object System.Collections.ArrayList
        $ADGroups.Add('GRP_StudentAccount') > $null
        $ADGroups.Add('GRP_StudentLicense') > $null
        #add other entities to OHHS
        if ('RUNNING START','IGRAD ACADEMY' -contains $Student.'Entity Name'){
            $ADGroups.Add('OHHS-Students') > $null
        }

        $School=($StudentOUTable[$Student.'Entity Name'] -split 'OU=')[1].Replace(',',$null)
        
        if ($School -ne 'Out of District'){
            $StuSchoolGroup = "$School-Students"
            $ADGroups.Add($StuSchoolGroup) > $null
        }
        if (!$NoGradYearStu){
            $StuGradGroup = "$School-$($Student.'Stu Grad Yr')"
            $ADGroups.Add($StuGradGroup) > $null
        }else{
            $StuGradGroup = $null
        }

        if ($School -ne 'Out of District'){
            #Check if groups are present and if not create
            #Student Year Group
            if (($StuGradGroup) -and (![adsi]::Exists("LDAP://$DCServer/CN=$StuGradGroup,OU=$($Student.'Stu Grad Yr'),$($StudentOUTable[$Student.'Entity Name'])$StudentRootOU"))){
                #Group does not exist create
                New-ADGroup -Name $StuGradGroup -groupscope Global -path "OU=$($Student.'Stu Grad Yr'),$($StudentOUTable[$Student.'Entity Name'])$StudentRootOU" `
                    -OtherAttributes @{mail="$StuGradGroup@students.$Domain"} `
                    -Server $DCServer
            }
            #Student School Group
            if (! [adsi]::Exists("LDAP://$DCServer/CN=$StuSchoolGroup,$($StudentOUTable[$Student.'Entity Name'])$StudentRootOU")){
                #Group does not exist create
                New-ADGroup -Name $StuSchoolGroup -groupscope Global -path "$($StudentOUTable[$Student.'Entity Name'])$StudentRootOU" `
                    -OtherAttributes @{mail="$StuSchoolGroup@students.$Domain"} `
                    -Server $DCServer
            }
        }


        #Check for student
        $ID = [string]$Student.'Other ID'

        #Find student in AD table
        $ADStudent = $null
        foreach ($s in $ADStudents){
            if ($s.EmployeeID -eq $ID -or $s.SamAccountName -eq $SamAccount){
                 $ADStudent = $s
                 break
            }
        }
        
        #skip if student is set to not track
        if (!($ADStudent.Memberof -like '*Grp_DoNotTrack*')){
            if ($null -ne $ADStudent){

                #Verify account is not disabled
                if ($ADStudent.Enabled -eq $false){
                    Try{
                        Set-ADUser -Identity $ADStudent.ObjectGUID -Enabled $True -ErrorAction Stop
                    }catch{
                        $ErrorMessage = "Unable to reenable $($ADStudent.SamAccountName)"
                        Write-Host ("$ErrorMessage" + ": $($Error[0])") -ForegroundColor Red
                        Write-Output $ErrorMessage | Out-File $Log -Append
                        $Error[0] | Out-File $Log -Append
                        $ErrorEmail = $ErrorEmail + "<p>$ErrorMessage</p>" + "<pre>$($Error[0])</pre>"
                    }
                }

                #Remove grace window if found
                if ($ADStudent.Description -like 'To be Disabled:*'){
                    Set-ADUser -Identity $ADStudent.ObjectGUID -Description $null -Server $DCServer
                }
            
                #Verify student name with records
                if (($ADStudent.GivenName -ne $Student.'Stu First Name') -or ($ADStudent.Surname -ne $Student.'Stu Last Name')){
                    Set-ADUser -Identity $ADStudent.ObjectGUID -DisplayName "$(($Student.'Stu First Name').ToUpper()) $(($Student.'Stu Last Name').ToUpper())" `
                                            -SamAccountName $SamAccount `
                                            -GivenName $(($Student.'Stu First Name').ToUpper()) `
                                            -Surname $(($Student.'Stu Last Name').ToUpper()) `
                                            -UserPrincipalName ($Username + '@' + 'students.' + $Domain) `
                                            -EmailAddress $Email `
                                            -Server $DCServer
                    Rename-ADObject -Identity $ADStudent.ObjectGUID -NewName $Username -Server $DCServer
                
                #Verify student email matches
                }elseif ($ADStudent.EmailAddress -ne $Email){
                    Set-ADUser -Identity $ADStudent.ObjectGUID -EmailAddress $Email -Server $DCServer
                    Rename-ADObject -Identity $ADStudent.ObjectGUID -NewName $Username -Server $DCServer
                }

                #Verify Other ID
                if ($ADStudent.EmployeeID -ne $Student.'Other ID'){
                    Set-ADUser -Identity $ADStudent.ObjectGUID -EmployeeID $Student.'Other ID' -Server $DCServer
                }

                #Verify OU
                $ADStuOU = $ADStudent.DistinguishedName -replace '^CN=.+?,'
                if ($ADStuOU -ne $ADOU){
                    Move-ADObject -Identity $ADStudent.ObjectGUID -TargetPath $ADOU -Server $DCServer
                }
                #Verify Groups if Auditing
                #check for missing groups
                foreach ($Group in $ADGroups){                       
                    if (!($ADStudent.MemberOf -like "*$Group*")){
                        #Missing default group
                        Add-ADGroupMember -Identity $Group -Members $ADStudent.ObjectGUID -Server $DCServer
                        $Message = "Missing group $Group added for $Username"
                        Write-Warning $Message
                        Write-Output $Message | Out-File $Log -Append
                    }
                }
                #check for excessive groups
                foreach ($ADStudentGroup in $ADStudent.MemberOf){
                    #convert DN to just group name
                    $ADStudentGroup = $ADStudentGroup -replace '^CN=|,.*$'
                    if (!($ADGroups -like $ADStudentGroup)){
                        #Remove group
                        Remove-ADGroupMember -Identity $ADStudentGroup -Members $ADStudent.ObjectGUID -Server $DCServer -Confirm:$False
                        $Message = "Removed non default group $ADStudentGroup for $Username"
                        Write-Warning $Message
                        Write-Output $Message | Out-File $Log -Append
                    }
                }
            }else{
                #Student does not exist in AD, create
                Try {
                    New-ADUser -Name $Username `
                               -DisplayName "$(($Student.'Stu First Name').ToUpper()) $(($Student.'Stu Last Name').ToUpper())" `
                               -Path $ADOU `
                               -AccountPassword (ConvertTo-SecureString $PWD -AsPlainText -Force) `
                               -Enabled 1 `
                               -GivenName $(($Student.'Stu First Name').ToUpper()) `
                               -Surname $(($Student.'Stu Last Name').ToUpper()) `
                               -EmployeeID $($Student.'Other ID') `
                               -UserPrincipalName ($Username + '@' + 'students.' + $Domain) `
                               -SamAccountName $SamAccount `
                               -EmailAddress $Email `
                               -CannotChangePassword $true `
                               -Server $DCServer
                }catch{
                    $ErrorMessage = "Error with student $Username, ID: $($Student.'Other ID')"
                        if ($Error[0] -like '*is not unique forest-wide*'){
                            #possible duplicate because Other ID changed
                            $ADDuplicate = $ADStudents | where {$_.SamAccountName -eq $SamAccount}
                            if ($ADDuplicate){
                                $ErrorMessage = $ErrorMessage + "`nAn account was found with the same alphakey but using a different other ID.`nOld ID $($ADDuplicate.EmployeeID)." `
                                                              + "`nVerify and update AD account (AD employee ID attribute) to the correct `"other ID`" to resolve conflict."
                            }
                        }
                
                    Write-Host $ErrorMessage -ForegroundColor Red
                    Write-Output $ErrorMessage | Out-File $Log -Append
                    $Error[0] | Out-File $Log -Append
                    $ErrorEmail = $ErrorEmail + "<p>$ErrorMessage</p>" + "<pre>$($Error[0])</pre>"
                    Continue
                }
                $NewStudent = New-Object PSObject
                $NewStudent | Add-Member NoteProperty NAME "$(($Student.'Stu First Name').ToUpper()) $(($Student.'Stu Last Name').ToUpper())"
                $NewStudent | Add-Member NoteProperty SAM  $SamAccount
                $NewStudent | Add-Member NoteProperty PWD  $PWD
                $ParentEmails = [System.Collections.ArrayList]@()
                        if ($Student.'F1/G1 Email'){
                $ParentEmails.Add($Student.'F1/G1 Email') > $null
            }
                        if ($Student.'F1/G2 Email'){
                $ParentEmails.Add($Student.'F1/G2 Email') > $null
            }
                $ParentEmails = $ParentEmails -join ','
                $NewStudent | Add-Member NoteProperty PEMAILS $ParentEmails
                $NewStudents.Add($NewStudent) > $null

                #Add to student group
                foreach ($ADGroup in $ADGroups){
                    Add-ADGroupMember -Identity $ADGroup -Members $SamAccount -Server $DCServer
                }
            }
        }
        if ($Progress){
            $i++
        }
    }else{
        $ErrorMessage = "No entity exists for $Username with entry $($Student.'Entity Name'). Check and update hashtable in script."
        Write-Warning $ErrorMessage
        Write-Output $ErrorMessage | Out-File $Log -Append
        $ErrorEmail = $ErrorEmail + "<p>$ErrorMessage</p>"
    }

}

if ($Progress){
    Write-Progress -Activity 'Processing Skyward Data' -Completed
}

#check for users no longer in AD
Write-Verbose 'Checking for orphaned students.' -Verbose
$ActiveADStudents = Get-AdUser -LDAPFilter "(&(objectCategory=user)(memberof=$StudentGroupDN)(!userAccountControl:1.2.840.113556.1.4.803:=2)(!(memberof=$DoNotTrackGroupDN)))" `
                         -Properties Name,EmailAddress,DistinguishedName,Description,EmployeeID,Enabled,MemberOf,ObjectGUID `
                         -Server $DCServer
$OrphanedStudents = Compare-Object -ReferenceObject $SkywardStudents.'Other ID' -DifferenceObject $ActiveADStudents.EmployeeID |`
                    Where {$_.SideIndicator -eq '=>'} |`
                    Select InputObject

$i = 1
$TotalStudents = $OrphanedStudents.InputObject.Count

foreach ($OStudentID in $OrphanedStudents.InputObject){

    if ($Progress){
        #Progress bar
        Write-Progress -Activity 'Processing Orphaned Students' -Status "Student $i of $TotalStudents" -PercentComplete ($i / $TotalStudents * 100)
    }

    #Variables
    $GradStudent = $null

    #find orphaned student in table
    $OADStudent = $null
    foreach ($s in $ActiveADStudents){
        if ($s.EmployeeID -eq $OStudentID){
                $OADStudent = $s
                break
        }
    }

    $OADSOU = $OADStudent.DistinguishedName -replace '^CN=.+?,'

    #pull grad year from OU
    $GradYear = (($OADStudent.DistinguishedName -split ',')[1] -split 'OU=')[1]
    
    #Archive OU Prep
    if ($GradYear -le $CurrentYear){
        $DisableOU = 'OU=zzGraduates,'
    }else{
        $DisableOU = 'OU=Inactive Students,'
    }
    $ArchiveOU = "OU=$CurrentYear," + $DisableOU + $StudentRootOU
    #Check/create archive year OUs
    if (! [adsi]::Exists("LDAP://$ArchiveOU")){
        #OU does not exist create
        New-ADOrganizationalUnit -Name $CurrentYear -Path ($DisableOU + $StudentRootOU) `
            -Server $DCServer
    }
        
    ########################   NORMAL GRADUATES   ########################
    #Verify grad year with date
    if ($GradYear -match '^\d+$'){
        if ($GradYear -eq $CurrentYear){
            if ($CurrentMonth -in 6..8){
                $GradStudent = $true
                if (($OADSOU -split ',')[1] -ne 'OU=zzGraduates'){
                    #First detection of student graduation. Move account and notify student
                    $MailBody = @"
                        <style>
                            body,p,h3 { font-family: calibri; }
                            h3  { margin-bottom: 5px; }
                            th  { text-align: center; background: #003829; color: #FFF; padding: 5px; }
                            td  { padding: 5px 20px; }
                            tr  { background: #E7FFF9; }
                        </style>

                        <h3>Account Closure Notice</h3>
			            <p>This is an automated message notifying you that your account will be closed by the beginning of September.</p>
                        <p>There will be another message to remind you two weeks before closing your account</p>
                        <p>If you have any data you want to keep (emails, google docs) please see <a href="https://support.google.com/accounts/answer/6386856" target="_blank">this support page</a>.</p>
"@
                    Send-MailMessage -SmtpServer 'smtp-relay.gmail.com' -To "$($OADStudent.EmailAddress)" -From 'noreply@domain.net' -Subject 'District Account Closure Notice' -Body $MailBody -BodyAsHtml

                    #Remove license
                    Remove-ADGroupMember -Identity 'GRP_StudentLicense' -Members $OADStudent.ObjectGUID -Server $DCServer -Confirm:$False -ErrorAction Ignore

                    #move OU
                    Move-ADObject -Identity $OADStudent.ObjectGUID -TargetPath $ArchiveOU -Server $DCServer
                }

            }
        }
    }
    #Disable account and move to inactive
    if (!$GradStudent){
        #Two week grace window before disable
        $GraceWindow = ($OADStudent.Description -split 'To be Disabled: ')[1]
        if ($GraceWindow){
            try {
                $GraceWindow = $GraceWindow | Get-Date -ErrorAction Stop
            }catch{
                $GraceWindow = $null
            }
        }
        if (!$GraceWindow){
            #Grace period not set. Timestamp, and continue
            Try {
                $OADStudent | Set-ADUser -Description "To be Disabled: $GraceDate" -Server $DCServer -ErrorAction Stop
            }catch{
                $ErrorMessage = "Error occured while prepping account for grace period for $($OADStudent.Name) in AD."
                Write-Warning $ErrorMessage
                Write-Output $ErrorMessage | Out-File $Log -Append
                $Error[0] | Out-File $Log -Append
                $ErrorEmail = $ErrorEmail + "<p>$ErrorMessage</p>" + "<pre>$($Error[0])</pre>"
                continue
            }

            #Email grads giving two week notice
            if (($OADSOU -split ',')[1] -eq 'OU=zzGraduates'){
                $MailBody = @"
                    <style>
                        body,p,h3 { font-family: calibri; }
                        h3  { margin-bottom: 5px; }
                        th  { text-align: center; background: #003829; color: #FFF; padding: 5px; }
                        td  { padding: 5px 20px; }
                        tr  { background: #E7FFF9; }
                    </style>

                    <h3>Account Closure Notice</h3>
			        <p>This is an automated message notifying you that your account will be closed in two weeks.</p>
                    <p>This is the final warning before closure.</p>
                    <p>If you have any data you want to keep (emails, google docs) please see <a href="https://support.google.com/accounts/answer/6386856" target="_blank">this support page</a>.</p>
"@
                Send-MailMessage -SmtpServer 'smtp-relay.gmail.com' -To "$($OADStudent.EmailAddress)" -From 'noreply@domain.net' -Subject 'District Account Closure Two Week Notice' -Body $MailBody -BodyAsHtml
            }

            ########################   EARLY GRADUATES   ########################

            #If graduate of this year, send two week notice and move to grad archive OU
            if ($GradYear -eq $CurrentYear){
                $MailBody = @"
                    <style>
                        body,p,h3 { font-family: calibri; }
                        h3  { margin-bottom: 5px; }
                        th  { text-align: center; background: #003829; color: #FFF; padding: 5px; }
                        td  { padding: 5px 20px; }
                        tr  { background: #E7FFF9; }
                    </style>

                    <h3>Account Two Week Notice</h3>
			        <p>This is an automated message notifying you that your account is scheduled to be closed in two weeks due to data received from registrar.</p>
                    <p>If this is is incorrect, please contact the front office of your school to check your student registration.</p>
                    <p>The scheduled date for closure is $GraceDate</p>
                    <p>As a reminder, if you have any data you want to keep (emails, google docs) please see <a href="https://support.google.com/accounts/answer/6386856" target="_blank">this support page</a>.</p>
"@
                Send-MailMessage -SmtpServer 'smtp-relay.gmail.com' -To "$($OADStudent.EmailAddress)" -From 'noreply@domain.net' -Subject 'District Account Closure Two week Notice' -Body $MailBody -BodyAsHtml
                #move to archive OU if not already there
                if ($OADSOU -ne $ArchiveOU){
                    Move-ADObject -Identity $OADStudent.ObjectGUID -TargetPath $ArchiveOU -Server $DCServer
                }

            }

            $GraceStudents.Add($OADStudent.Name) > $null
            Continue
        }else{
            #within grace window, skip disable for now
            if ($GraceWindow -gt $Today){
                Continue
            }
        }

        ####################    DISABLE ACCOUNT    ####################
        #Remove groups
        foreach ($OldGroup in $OADStudent.MemberOf){
            if ($OldGroup -notlike 'CN=GRP_StudentAccount,*'){
                Remove-ADGroupMember -Identity $OldGroup -Members $OADStudent.ObjectGUID -Server $DCServer -Confirm:$False
            }
        }

        #add account to array to keep track and disable
        $DisabledStudents.Add($OADStudent.Name) > $null
        $OADStudent | Disable-ADAccount -Server $DCServer

        #randomize password
        $RandomPassword = [string] -join ((33..126) | Get-Random -Count 12 | % {[char]$_})

        if ($RandomPassword){
            Set-ADAccountPassword -Identity $OADStudent.ObjectGUID -NewPassword (ConvertTo-SecureString $RandomPassword -AsPlainText -Force) -Server $DCServer
        }

        #move to archive OU if not already there
        if ($OADSOU -ne $ArchiveOU){
            Move-ADObject -Identity $OADStudent.ObjectGUID -TargetPath $ArchiveOU -Server $DCServer
        }
    }
    $i++
}
Write-Progress -Activity 'Checking for Orphaned Students' -Completed

#Sync changes with Google via GCDS
Write-Verbose 'Running GCDS Student sync' -Verbose
if ($env:COMPUTERNAME -eq $GCDSServer){
    &'\\domain.net\dfs\AdminScripts\Scheduled-Tasks\Start-GCDS.ps1'
}else{
    #More verbose script if not running on server
    &'\\domain.net\dfs\AdminScripts\Tools\Start-ADGoogleSync.ps1' -RunOnce
}

#Wait a few minutes to allow google to create accounts from sync
if ($NewStudents){
    Start-Sleep -s $WaitTime
}

#Reset Passwords again so sync can be captured.
Write-Verbose 'Processing password and emails for new students.' -Verbose
$i = 0
$TotalStudents = $NewStudents.Count
foreach ($NewStudent in $NewStudents){
    if ($Progress){
        #Progress bar
        Write-Progress -Activity 'Syncing Passwords' -Status "Student $i of $TotalStudents" -PercentComplete ($i / $TotalStudents * 100)
    }
    Set-ADAccountPassword -Identity $NewStudent.SAM -NewPassword (ConvertTo-SecureString $NewStudent.PWD -AsPlainText -Force) -Server $DCServer

    #Send out welcome email
    $ParentEmails = $NewStudent.PEMAILS -split ','
    $ParentEmails = $ParentEmails | select -Unique
    
    if (!$NoEmail){
        Foreach ($ParentEmail in $ParentEmails){
            if ($ParentEmail){
                #E-mail template
                $MailBody = @"
                    <style>
                        body,p,h3 { font-family: calibri; }
                        h3  { margin-bottom: 5px; }
                        th  { text-align: center; background: #003829; color: #FFF; padding: 5px; }
                        td  { padding: 5px 20px; }
                        tr  { background: #E7FFF9; }
                    </style>

                    <h3>Welcome to Public Schools</h3>

                    <p>An account has been created for your student to access their Public Schools email and other educational digital resources.</p>
                    <p>The following information will be needed for them to log into their district-issued Google account:</p>
                    <p><b>Username:</b> $($NewStudent.SAM)@students.domain.net</p>
                    <p><b>Password:</b> $($NewStudent.PWD)</p>
                    <p>To login to your new account, please visit www.gmail.com and enter the username and password from above.</p>

                    <p>If you have found any typos in your child's name, please contact your student's school office to make the appropriate changes.</p>
"@
                Send-MailMessage -SmtpServer 'smtp-relay.gmail.com' -To $ParentEmail -From 'noreply@domain.net' -Subject 'Disctrict Account Creation' -Body $MailBody -BodyAsHtml
            }
        }
    }

    $i++
}

#Send email if errors occured
if (($ErrorEmail) -and ($env:USERNAME -eq 'adautomation') -and (!$NoEmail)){
    Write-ErrorEmail -ErrorString $ErrorEmail
}

Write-Progress -Activity 'Syncing Passwords' -Completed

#Finish message
$EndTime = Get-Date
$Message = "Script completed at $EndTime. Script took $(New-TimeSpan -Start $StartTime -End $EndTime)" +`
           "`n$($NewStudents.Count) new students were created, $($GraceStudents.Count) put in grace period, and $($DisabledStudents.Count) were disabled."

if ($NewStudents.Count -gt 0){
    $Message += "`n`nNew Accounts:`n" + ($NewStudents.SAM | Out-String)

}
if ($GraceStudents.Count -gt 0){
    $Message += "`n`nAccounts in Grace Period:`n" + ($GraceStudents | Out-String)
}
if ($DisabledStudents.Count -gt 0){
    $Message += "`n`nDisabled Accounts:`n" + ($DisabledStudents | Out-String)
}

Write-Host $Message -ForegroundColor Cyan
Write-Output $Message | Out-File $Log -Append

#Log cleanup
$MaxLogLength = 301
$LogText = (Get-Content $Log | Out-String) -split '\*{20}'
if ($LogText.Count -gt $MaxLogLength){
    $Output = [System.Text.StringBuilder]''
    $i=$LogText.Count
    foreach ($Line in $LogText){
        #skip entry if above max length
        if ($i -lt $MaxLogLength){
            if ($Line -ne ''){
                if ($Line -like ' Script Started by *'){
                    $Output.AppendLine('*'*20 + $Line + '*'*20) > $null
                }else{
                    $Output.AppendLine($Line.Trim()) > $null
                    $Output.AppendLine() > $null
                }
            }
        }
        $i--
    }
    $Output.ToString() | Out-File $Log
}
