Option Explicit

Dim objShell, objExec, strUserName, strOutput, intLine, nulrtrnsuccess
Dim intPosition, strIP, arrEntrySplit, strComputerName, intRemote, strWinsServ

'Set WINS Server Name here:
strWinsServ = "hbseaadcp02"

Set objShell = CreateObject("WScript.Shell")
strUserName = InputBox("Please enter a username and press [Enter]:" & vbCRLF & vbCRLF & _
   "(Or click [Cancel] to exit.)", "Find That User's Computer!", "")

Do While strUserName <> ""
'EDIT THE FOLLOWING LINE TO INCLUDE YOUR WINS SERVERNAME OR IP ADDRESS
   Set objExec = objShell.Exec("NETSH WINS SERVER \\" & strWinsServ & " SHOW NAME " & strUserName & " 03")
   For intLine = 1 To 4
      objExec.StdOut.ReadLine
   Next
   If objExec.StdOut.ReadLine <> "The name does not exist in the WINS database." Then
      For intLine = 1 To 5
         objExec.StdOut.ReadLine
      Next
      strOutput = objExec.StdOut.ReadLine
      If Left(strOutput,10) = "IP Address" Then
         intPosition = InStr(strOutput, ":")
         strIP = Right(strOutput, (Len(strOutput) - intPosition) - 1)
         Set objExec = objShell.Exec("NBTSTAT -A " & strIP)
         Do While Not objExec.StdOut.AtEndOfStream
            strOutput = objExec.StdOut.ReadLine
            If InStr(strOutput,"Host not found") <> 0 Then
                 MsgBox "The username '" & strUserName & "' was found in the WINS database. However, the associated workstation at IP address '" & strIP & "' no longer appears to be on the network.",,"NOT FOUND - Find That User's Computer!"
            ElseIf InStr(strOutput,"<20>") <> 0 Then
               arrEntrySplit = Split(strOutput)
'If UCase(arrEntrySplit(4)) <> UCase(strUserName) And _
'If UCase(arrEntrySplit(4)) <> UCase(strComputerName) And _
'UCase(arrEntrySplit(4)) <> strComputerName & "$" Then
               strComputerName = UCase(arrEntrySplit(4))
               If InStr(strComputerName,"CITRIX") <> 0 Then
                    MsgBox "Sorry!" & vbCRLF & vbCRLF & "The user appears to have last logged onto a Citrix environment, and therefore cannot be pinpointed on a workstation.",, "Find That User's Computer!"
               ElseIf strComputerName <> "" Then
                    nulrtrnsuccess = InputBox("The requested user '" & strUserName & _
                    "' is logged on to: ","SUCCESS - Find That User's Computer!",strComputerName)
               Else
                    MsgBox "Sorry!" & vbCRLF & vbCRLF & "The username '" & strUserName & "' was found at IP Address '" & strIP & "'. However, there was an error when trying to find the computername for that IP Address." _
                    & vbCRLF & "Either use this IP Address to find the workstation or click OK and try to find the user again.",, _
                    "ERROR - Find That User's Computer!"
               End If
'End If
            End If
         Loop
      Else
           MsgBox "Sorry!" & vbCRLF & vbCRLF & "An error has occured while querying the WINS database. The user may be listed but may not have logged on from within your local network recently." _
           & vbCRLF & "You can also make sure you entered the WINS servername into this script correctly and that you have rights to query the WINS database.",, _
           "ERROR - Find That User's Computer!"
      End If
   Else
      MsgBox "You were able to access the WINS database. However, the username '" & strUserName & "' was not found." _
         & vbCRLF & "If that username is incorrect, click OK and try again. Otherwise, the user does not appear to be logged onto the network at this time.",,"USER NOT FOUND - Find That User's Computer!"
   End If
   strComputerName = ""
   strUserName = InputBox("Please enter a username and press [Enter]:" & vbCRLF & vbCRLF & _
      "(Or click [Cancel] to exit.)", "Find That User's Computer!", "")
Loop




