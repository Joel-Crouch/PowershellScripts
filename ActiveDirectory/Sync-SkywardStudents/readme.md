## Instructions

Make sure to adjust default parameters and adjust the StudentOUTable hash table to match your skyward entries and OU structure.

Account that is running script will need write permissions to the student OU and children objects. The script will create OUs for grad years if they do not exist.

A default student group is used for tracking which accounts are created and maintained by the script. This script uses GRP_StudentAccount.

Groups are used to track licensing with GCDS. This script uses GRP_StudentLicense 
for tracking student license.

Adjust email recipients for error emails (line 89)

WinSCP is used to pull export from FTP server. Account used for automated script has read only access for security.

Lines 152 to 157 will need to be adjusted for FTP server and account. WinSCP can generate this for you under Session/Generate URL/Code

Additional skyward entities were used to include with the high school OU. These can be removed or modified on line 253.

Default student password will need to be adjusted in line 343. In this example, the student password is the string “default” + their lunch pin (e.g. default19293).

Graduate OUs may need to be adjusted in line 440, 442

Email details will need to be adjusted in line 460, 522, and 544

The scheduled task that runs will need to be adjusted in line 587 and 590

The script referenced in that line is a simple script used to launch GCDS with logs
```
$LogDir = "\\domain.net\dfs\AdminScripts\AD\Sync-SkywardStudents"

."C:\Program Files\Google Cloud Directory Sync\sync-cmd.exe" -a `
    -c "C:\Users\adskywardsync\StudentConfig.xml" `
    -r "$LogDir\GCDSSyncChanges.txt" `
    *> "$LogDir\GCDSSyncOutput.txt"
```

Parent email starting on line 614 will need to be adjusted.

If you have a student that is in skyward but you do not want the script to move the student or make modifications to the account to match what is in skyward, you can put “do not track” in the description field of the student and the script will skip over that student until the description is removed.
