Option Explicit

Dim WshShell, fso, currentDirectory, powershellPath, scriptPath, command

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get the directory where this VBS file is located
currentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)

' Full path to the PowerShell script
scriptPath = currentDirectory & "\SpreadsheetWrangler.ps1"

' Make sure the script exists
If Not fso.FileExists(scriptPath) Then
    MsgBox "Error: Cannot find SpreadsheetWrangler.ps1 in the same directory as this launcher.", vbCritical, "Spreadsheet Wrangler Launcher"
    WScript.Quit
End If

' PowerShell executable path
powershellPath = "powershell.exe"

' Build the command
command = powershellPath & " -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & scriptPath & """"

' Run the command with window hidden (0) and don't wait for it to complete (False)
WshShell.Run command, 0, False

' Clean up
Set WshShell = Nothing
Set fso = Nothing
