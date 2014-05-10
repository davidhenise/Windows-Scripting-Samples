Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "Example CDO Message" 
objMessage.From = "deskeng@homestreet.com" 
objMessage.To = "deskeng@homestreet.com" 
objMessage.TextBody = "This is some sample message text."

'==This section provides the configuration information for the remote SMTP server.
'==Normally you will only change the server name or IP.
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

'Name or IP of Remote SMTP Server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "homail01"

'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 

objMessage.Configuration.Fields.Update

'==End remote SMTP server configuration section==

objMessage.Send
