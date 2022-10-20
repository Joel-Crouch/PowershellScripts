set shell = CreateObject("WScript.Shell")
AppData = shell.ExpandEnvironmentStrings("%APPDATA%")
command = "powershell.exe -file " & AppData & "\AdminScripts\Backup-UserFiles.ps1"

shell.Run command,0
'WScript.Echo command