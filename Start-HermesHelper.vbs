Option Explicit

Dim fso
Dim shell
Dim scriptDir
Dim scriptPath
Dim command

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
scriptPath = fso.BuildPath(scriptDir, "HermesHelper.ps1")

If Not fso.FileExists(scriptPath) Then
    MsgBox "HermesHelper.ps1 was not found next to this launcher." & vbCrLf & _
        "Expected: " & scriptPath, vbCritical, "Hermes Helper"
    WScript.Quit 1
End If

command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & QuoteArg(scriptPath)

shell.Run command, 0, False

Function QuoteArg(value)
    QuoteArg = """" & Replace(value, """", """""") & """"
End Function
