[CmdletBinding()]
param (
    $RegPath = "HKCU:\Software\AdminScripts\IT-UserBackup",
    $GDBackupDir = "IT-Backup"
)

#Find path of google drive
$GDrive = Get-PSDrive | Where-Object {$_.Description -like 'Google Drive*'}
$GDPath = "$($GDrive.Root)My Drive\$GDBackupDir"

#Terminate if google drive is not available
if (!($GDrive) -or !(Test-Path "$($GDrive.Root)My Drive" -ErrorAction SilentlyContinue)){
    Write-Warning "Google file stream is not currently accessible."

    #Check if GDrive has been configured
    if (!(test-path "$env:LOCALAPPDATA\Google\DriveFS\global_feature_config")){
        #check if drive has ever been run before
        if (!(test-path $RegPath)){
            #warn user with popup
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            $Message = 'Google Drive is currently not set up on this computer.'+
                       "`n`nThis software is needed to backup local documents."+
                       "`n`nWould you like to setup Google Drive?"+
                       "`n`nHit `"No`" to never be prompted again or `"Cancel`" to be reminded later."
            $MsgBoxInput =  [System.Windows.MessageBox]::Show($Message,'Google Drive Backup','YesNoCancel','Warning')
            switch  ($MsgBoxInput) {
                'Yes' {
                    #Old code. Use shortcut to simplify
                    <#Find latest google drive executable and start process
                    $CurrentDir = Get-ChildItem -Path "C:\Program Files\Google\Drive File Stream" -Directory | `
                                  Sort-Object Name -Descending | `
                                  Where-Object {$_.Name -notlike "Drivers"} | `
                                  Select-Object -First 1
                    $LaunchExe = "$($CurrentDir.FullName)\GoogleDriveFS.exe"#>
                    $GDriveShortcut = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Google Drive.lnk"
                    if (Test-Path $GDriveShortcut){
                        start-process $GDriveShortcut
                        #Start loop waiting for google drive access
                        $Timer = 0
                        Do {
                            $GDrive = Get-PSDrive | Where-Object {$_.Description -like 'Google Drive*'}
                            Start-Sleep -s 1
                            $Timer++
                        }until(($GDrive) -or ($Timer -gt 600))
                        if ($Timer -lt 600){
                            #rerun script
                            &$PSCommandPath
                        }
                    }
                }'No' {
                    #Write key to ignore
                    Write-Verbose "Setting key to never display message again" -Verbose
                    New-Item -Path $RegPath -Force | Out-Null
                }
            }
        }
    }       
    exit
}

#Check if folder exists
if (!(Test-Path $GDPath)){
    New-Item $GDPath -ItemType Directory | Out-Null
    #hide folder on file stream
    $F = Get-Item $GDPath
    $F.Attributes+="Hidden"
}

#Check if first run
if (!(Test-Path $RegPath)){
    #First run
    Write-Verbose "First run detected. Checking folders" -Verbose
    if (Test-Path "$GDPath\$Env:COMPUTERNAME"){
        #old backup found
        if (Test-Path "$GDPath\$Env:COMPUTERNAME-OLD"){
            #delete old folder
            Remove-Item "$GDPath\$Env:COMPUTERNAME-OLD" -Recurse -Force
        }
        #Keep old data incase of reimage and script is ran before data is copied off
        Write-Verbose "Old data found on first run. Keeping old data as OLD" -Verbose
        Rename-Item "$GDPath\$Env:COMPUTERNAME" "$GDPath\$Env:COMPUTERNAME-OLD"     
    }
    New-Item "$GDPath\$Env:COMPUTERNAME" -ItemType Directory | Out-Null
    New-Item -Path $RegPath -Force | Out-Null
}

########## Start Backup Process ############

#Grab library paths
$DocumentsLibrary = ([Environment]::GetFolderPath("MyDocuments"))
$DesktopLibrary = ([Environment]::GetFolderPath("Desktop"))

#Skip backup if libraries have been changed to google drive
if ($DocumentsLibrary -notlike "$($GDrive.Root)*"){
    robocopy $DocumentsLibrary "$GDPath\$Env:COMPUTERNAME\Documents" /MIR /FFT /Z /XJF /XJD /XA:H /R:3 /W:60 /MT:4
}elseif(Test-Path "$GDPath\$Env:COMPUTERNAME\Documents"){
    #Clean up unnecessary backup
    Remove-Item "$GDPath\$Env:COMPUTERNAME\Documents" -Recurse -Force
}
if ($DesktopLibrary -notlike "$($GDrive.Root)*"){
    robocopy $DesktopLibrary "$GDPath\$Env:COMPUTERNAME\Desktop" /MIR /FFT /Z /XJF /XJD /XA:H /R:3 /W:60 /MT:4
}elseif(Test-Path "$GDPath\$Env:COMPUTERNAME\Desktop"){
    #Clean up unnecessary backup
    Remove-Item "$GDPath\$Env:COMPUTERNAME\Desktop" -Recurse -Force
}

#Find Bookmarks
$ChromeBookmarks = Get-ChildItem -Path "$env:LOCALAPPDATA\Google\Chrome\User Data" -Filter Bookmarks -Recurse -ErrorAction SilentlyContinue -Force |`
                                 where {$_.Directory -notlike "*SnapShots*"}
Foreach ($Bookmark in $ChromeBookmarks){
    $PDir = ($Bookmark.Directory -split "\\")[-1]
    $Source = $Bookmark.Directory
    $Destination = "$GDPath\$Env:COMPUTERNAME\Chrome\$PDir"
    $File = $Bookmark.Name
    robocopy $Source $Destination $File /R:3 /W:60
}
