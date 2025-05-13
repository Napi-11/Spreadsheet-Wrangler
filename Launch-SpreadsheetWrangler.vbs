Set WshShell = CreateObject("WScript.Shell")
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
strCommand = "powershell.exe -ExecutionPolicy Bypass -File """ & strPath & "\SpreadsheetWrangler.ps1"""
WshShell.Run strCommand, 0, False
