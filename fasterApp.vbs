' For advanced users.
' Does the same thing as app.vbs with fewer if's checks and buts
' This will only save you 1ms, so use only if speed is mission cricital
Const RULES_CONFIG_FILE = "URIconfigs.ini"
Const LINE_SEPARATOR = "_##_"
Set objShell = CreateObject("Wscript.Shell")
Set RulesConfig = CreateObject( "Scripting.FileSystemObject" ).OpenTextFile(Replace(WScript.ScriptFullName, WScript.ScriptName, RULES_CONFIG_FILE),1)
Dim arrSchemas
Set arrSchemas = CreateObject("System.Collections.ArrayList")
Do while not RulesConfig.AtEndOfStream
    arrSchemas.Add RulesConfig.ReadLine()
Loop
arrSchemas.Capacity = arrSchemas.Count
intRulesLength = arrSchemas.Capacity-1
ReDim arrRules(intRulesLength)
For i = 0 To intRulesLength
    arrRules(i) = Split(arrSchemas.Item(i), LINE_SEPARATOR)
Next
For i = 0 To intRulesLength
    arrSchemas.Insert i, arrRules(i)(0)
Next
incomingUrl = Split(WScript.Arguments(0), "://", 2)
itm = arrSchemas.IndexOf(incomingUrl(0), 0)
startCmd = Chr(34) & arrRules(itm)(1) & Chr(34) & _
            Space(1) & arrRules(itm)(2) & Space(1) & _
            Chr(34) & Trim(incomingUrl(1)) & Chr(34)
' MsgBox startCmd
objShell.Run startCmd, 1, False
WScript.Quit(0)