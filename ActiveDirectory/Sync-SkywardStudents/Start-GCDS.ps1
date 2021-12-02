[cmdletbinding()]
param (
    [switch]$Flush
)

if ($Flush){
    $FlushArgument = '-f'
}else{
    $FlushArgument = $null
}

$LogDir = "\\ohsd.net\dfs\AdminScripts\AD\Sync-SkywardStudents"

."C:\Program Files\Google Cloud Directory Sync\sync-cmd.exe" -a `
    $FlushArgument `
    -c "C:\Users\adskywardsync\StaffStudentConfig.xml" `
    -r "$LogDir\GCDSSyncChanges.txt" `
    *> "$LogDir\GCDSSyncOutput.txt"