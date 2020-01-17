requireAdmin

Const VBS_SCRIPT_APP = "app.vbs"
Const RULES_CONFIG_FILE = "URIconfigs.ini"
Const LINE_SEPARATOR = "_##_"

Set objShell = CreateObject("Wscript.Shell")

function applyURI(scheme)
    vbsPath = Chr(34) & Replace(WScript.ScriptFullName, WScript.ScriptName, VBS_SCRIPT_APP) & Chr(34)
    urlArg = Chr(34) & "%1" & Chr(34)
    objShell.RegWrite "HKCR\"&scheme&"\", ""
    objShell.RegWrite "HKCR\"&scheme&"\", "URL:"&scheme&" Protocol", "REG_SZ" 'Default value
    objShell.RegWrite "HKCR\"&scheme&"\URL Protocol", "", "REG_SZ"
    objShell.RegWrite "HKCR\"&scheme&"\shell\", ""
    objShell.RegWrite "HKCR\"&scheme&"\shell\open\", ""
    objShell.RegWrite "HKCR\"&scheme&"\shell\open\command\", ""
    objShell.RegWrite "HKCR\"&scheme&"\shell\open\command\", "C:\Windows\System32\WScript.exe " & vbsPath & Space(1) & urlArg, "REG_SZ" 'Default value
end function

function readFile
    Dim i
    Dim intRulesLength
    Dim rulesFile
    
    rulesFile = Replace(WScript.ScriptFullName, WScript.ScriptName, RULES_CONFIG_FILE)

    Set arrSchemas = CreateObject("System.Collections.ArrayList")
    Set goFS = CreateObject( "Scripting.FileSystemObject" )
    Set RulesConfig = goFS.OpenTextFile(rulesFile,1)
    Do while not RulesConfig.AtEndOfStream
        arrSchemas.Add RulesConfig.ReadLine()
    Loop

    arrSchemas.Capacity = arrSchemas.Count
    intRulesLength = arrSchemas.Capacity-1

    if (intRulesLength < 0) Then log "Unable to read any rules"

    For i = 0 To intRulesLength
        applyURI(Split(arrSchemas.Item(i), LINE_SEPARATOR)(0))
    Next

    log "Success: All config file changes applied to the registry"
end function

readFile

Sub log (str)
    MsgBox str
    WScript.Quit(0)
End Sub

function requireAdmin
    If Not WScript.Arguments.Named.Exists("elevate") Then
    CreateObject("Shell.Application").ShellExecute WScript.FullName _
      , """" & WScript.ScriptFullName & """ /elevate", "", "runas", 1
    WScript.Quit
  End If
End function