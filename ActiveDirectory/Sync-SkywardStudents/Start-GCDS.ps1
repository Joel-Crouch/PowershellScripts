$LogDir = "\\domain.net\dfs\AdminScripts\AD\Sync-SkywardStudents"

."C:\Program Files\Google Cloud Directory Sync\sync-cmd.exe" -a `
    -c "C:\Users\adskywardsync\StudentConfig.xml" `
    -r "$LogDir\GCDSSyncChanges.txt" `
    *> "$LogDir\GCDSSyncOutput.txt"
