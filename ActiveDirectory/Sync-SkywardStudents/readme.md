## Instructions

Make sure to adjust default parameters and adjust the StudentOUTable hash table to match your skyward entries and OU structure.

Account that is running script will need write permissions to the student OU and children objects. The script will create OUs for grad years if they do not exist.

A default student group is used for tracking which accounts are created and maintained by the script. This script uses GRP_StudentAccount.

Groups are used to track licensing with GCDS. This script uses GRP_StudentLicense 
for tracking student license.

WinSCP is used to pull export from FTP server. Account used for automated script has read only access for security.

Lines 177 to 179 will need to be adjusted for FTP server and account. WinSCP can generate this for you under Session/Generate URL/Code

Default student password will need to be adjusted in line 233. In this example, the student password is the string “default” + their lunch pin (e.g. default19293).

Additional skyward entities were used to include with the high school OU. These can be removed or modified on line 273.

Graduate OUs may need to be adjusted starting in the blocks at 501

Email details will need to be adjusted starting at 535, 587, and 609

The scheduled task that runs will need to be adjusted starting in line 657

The script referenced in that line is a simple script used to launch GCDS with logs. It is included in this directory

Parent email starting on line 698 will need to be adjusted.

If you have a student that is in skyward but you do not want the script to move the student or make modifications to the account to match what is in skyward, you can add the account to a group (as specified in DoNotTrackGroupDN parameter) and the script will skip over that student until the description is removed.
