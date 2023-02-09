[cmdletbinding()]
param (
    [switch]$Flush,
	[string]$Config = 'C:\Users\Username\Documents\StaffStudentConfig.xml',
    [string]$LogDir = "\\domain\dfs\Logs"
)

if ($Flush){
    $FlushArgument = '-f'
}else{
    $FlushArgument = $null
}



."C:\Program Files\Google Cloud Directory Sync\sync-cmd.exe" -a `
    $FlushArgument `
    -c $Config `
    -r "$LogDir\GCDSSyncChanges.txt" `
    *> "$LogDir\GCDSSyncOutput.txt"
