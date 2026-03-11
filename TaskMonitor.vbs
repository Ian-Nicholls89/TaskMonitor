Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this VBS file is located
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
strTMScript = strScriptDir & "\TaskMonitor.ps1"

' Check if Task Monitor exists
If Not objFSO.FileExists(strTMScript) Then
    MsgBox "Error: TaskMonitor.ps1 not found in the same directory", 16, "Task Monitor"
    WScript.Quit 1
End If

' Run Task Monitor PowerShell script hidden (0 = hidden, no window)
objShell.Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & strTMScript & """", 0, False

WScript.Quit 0