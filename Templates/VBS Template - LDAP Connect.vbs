'-------------BEGIN LDAP CONNECTION-------------'
Const ADS_SEARCH_SCOPE_LEVELS = 3
Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
    '-------------BEGIN AUTHENTICATION-------------'
    'objConnection.Properties("User ID") = ""
    'objConnection.Properties("Password") = ""
    'objConnection.Properties("Encrypt Password") = TRUE
    'objConnection.Properties("ADSI Flag") = 1 
    '-------------END AUTHENTICATION-------------'
objConnection.Open "Active Directory Provider"
Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SEARCH_SCOPE_LEVELS
objCommand.Properties("Sort On") = "SN"
'-------------END LDAP CONNECTION-------------'

'-------------BEGIN QUERY-------------'
objCommand.CommandText = "SELECT * FROM 'LDAP://NOTES1' WHERE objectClass='dominoGroup'"
Set objRecordSet = objCommand.Execute
'-------------END QUERY-------------'
'-------------BEGIN ATTRIBUTE RETURN-------------'
objRecordSet.MoveFirst
Do Until objRecordSet.EOF
    On Error Resume Next
    Set objGroup = GetObject(objRecordSet.Fields("ADsPath").Value)
    If InStr (objGroup.mail,strSearch) Then msgbox Replace(objGroup.name,"CN=","")
    objRecordSet.MoveNext
Loop
'-------------END ATTRIBUTE RETURN-------------'
