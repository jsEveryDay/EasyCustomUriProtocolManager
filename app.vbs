Option Explicit
' On Error Resume Next

Dim objShell, goFS, RulesConfig, strArgsUrl, i
Dim arrSchemas, arrRules, intRulesLength, multiURL

Const RULES_CONFIG_FILE = "URIconfigs.ini"
Const LINE_SEPARATOR = "_##_"

Set objShell = CreateObject("Wscript.Shell")

' VBscipt has no way of taking in all arguments without looping through(1..23)
' This is just in case the user isn't using a browser (which encodes whitespace) or
' decides to pass multiple urls which Vbs interperets as separate arguments
' example: vlc://https://videoUrl.mp4 https://video1Url.mp4 htpp://youtube..
multiURL = WScript.Arguments.Count
If (multiURL = 0) Then 
    AlertBox "A URL/Argument is Required" + vbCrLf + "go back on Github and read the docs!", 1, 64
ElseIf (multiURL = 1) Then
    strArgsUrl = WScript.Arguments(0)
Else
    ReDim arr(WScript.Arguments.Count-1)
    For i = 0 To WScript.Arguments.Count - 1
        arr(i) = WScript.Arguments(i)
    Next
    strArgsUrl = Trim( Join(arr) )
End If

' Just in case Wscript.Shell fails
If (Err.Number <> 0) Then 
    AlertBox "Permissions or restriction issue for WScript", 2, 16
Else
    ' Proceed with command execution
    startCmd()
End If

function startCmd()
    ' Ex vlc://https://...etc
    Dim incomingUrl
    'Just sending the inital schema:// and making sure it exists on the config file
    incomingUrl = Split(strArgsUrl, "://", 2)
    ' Son now incomingurl(0) is vlc and incoming(1) is https://...etc
    i = boolMatchSchema(incomingUrl(0))
    if i > -1 Then
        startCmd = quoteIt(arrRules(i)(1)) & _
                    Space(1) & arrRules(i)(2) & Space(1) & _
                    quoteIt(Trim(incomingUrl(1)))
        'At this point, the whole command looks like this
        '"C:\Some Path With or Without Spaces\some.exe" arg --orarg /orlike this, any way works, as many as you want "URL"
        'Notice the binary is put in Quotes, then a space is added for the args then space "url" in quotes
        exec(startCmd)
    Else
        AlertBox "Your argument didn't match a URI, this came through: " + vbCrLf + vbCrLf+ strArgsUrl, 2, 48 
    End If          
end function

function boolMatchSchema(schema)
    Set goFS = CreateObject( "Scripting.FileSystemObject" )
    Set RulesConfig = goFS.OpenTextFile(getFullFilePath(RULES_CONFIG_FILE),1)
' Array class (stolen from .Net since VBScript needs a defined array(length))
' The alternative vb method is Array(-1) which takes a major performance hit
    Set arrSchemas = CreateObject("System.Collections.ArrayList")

    Do while not RulesConfig.AtEndOfStream
        arrSchemas.Add RulesConfig.ReadLine()
    Loop

    arrSchemas.Capacity = arrSchemas.Count
    intRulesLength = arrSchemas.Capacity-1

    if (intRulesLength < 0) Then AlertBox "Unable to read any rules", 2, 32

    ReDim arrRules(intRulesLength)
    For i = 0 To intRulesLength
        arrRules(i) = Split(arrSchemas.Item(i), LINE_SEPARATOR)
    Next
    'Makes no sense why this loops cant be joined in 1, somehow i goes ++ after the first exec
    'Behaves like an ES6 iterative or Java hasNext(), not in other scenarios, I think its the ArrList
    For i = 0 To intRulesLength
        arrSchemas.Insert i, arrRules(i)(0)
    Next
    'Respond back to startCmd with index id, so this can become private eventually
    boolMatchSchema = arrSchemas.IndexOf(schema, 0)
end function

Private Function quoteIt(ByVal strValue)
    quoteIt = Chr(34) & strValue & Chr(34)
End Function

private function getFullFilePath(fileName)
    Dim strFolder, fullFilePath
    strFolder = goFS.GetParentFolderName(WScript.ScriptFullName)
    fullFilePath = strFolder +"\"+ fileName
    If goFS.FileExists( fullFilePath ) Then 
        getFullFilePath = fullFilePath
    Else
        AlertBox fileName & " not found, please create it in the current script path.", 3, 48
    End If
end function

Sub AlertBox (someText, errorLevel, icon)
        ' Set objShell = CreateObject("Wscript.Shell")
        objShell.Popup someText ,10, "Oh sheeeeet", icon
        WScript.Quit(errorLevel)
        'just paranioa
        Set objShell = Nothing
End Sub

function exec(fullCommand)
        ' MsgBox fullCommand
        objShell.Run fullCommand, 1, False
        Set objShell = Nothing
        WScript.Quit(0)
end function