'Option Explicit

Dim objShell, objFSO, strTarget, strTargetList, objTextStream, RemotePCs, IE
Dim objWshScriptExec, objStdOut, strLine, Awake

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Function Browse()
Dim Q2, sRet, IE
On Error Resume Next
Q2 = chr(34)
Set IE = CreateObject("InternetExplorer.Application")
IE.visible = False
IE.Navigate("about:blank")
Do Until IE.ReadyState = 4
Loop
IE.Document.Write "<HTML><BODY><INPUT ID=" & _
Q2 & "Fil" & Q2 & "Type=" & Q2 & "file" & Q2 & _
"></BODY></HTML>"
With IE.Document.all.Fil
.focus
.click
Browse = .value
End With
IE.Quit
Set IE = Nothing
End Function

Function Ping(PC)
Set objWshScriptExec = objShell.Exec("ping.exe -n 1 " & PC)
Set objStdOut = objWshScriptExec.StdOut
awake=False
Do Until objStdOut.AtEndOfStream
    strLine = objStdOut.ReadLine
    awake = awake Or InStr(LCase(strLine), "bytes=") > 0
Loop
Ping = awake
End Function

Sub RunScript()
If Ping(strTarget) Then
    'DO SOMETHING
    outFile.WriteLine strTarget
End If
End Sub

Sub EndScript()
Set objShell=Nothing
Set objFSO=Nothing
Set IE=Nothing
Set objTextStream=Nothing
Wscript.Quit
End Sub

strTargetList = Browse()

Set objTextStream = objFSO.OpenTextFile(strTargetList)
RemotePCs = Split(objTextStream.ReadAll, vbCRLF)
objTextStream.Close
strTemp = WshShell.ExpandEnvironmentStrings("%temp%") & "\"
strOutFile = strTemp & "Action Successful on These Targets.log"
Set outFile = objFSO.CreateTextFile(strOutFile)

For Each strTarget in RemotePCs
    RunScript()
Next

outFile.Close
WshShell.Run("notepad " & strOutFile)

EndScript()
