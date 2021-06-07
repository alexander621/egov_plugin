<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkduplicatecitizens.asp
' AUTHOR: Steve Loar
' CREATED: 10/22/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed citizen is not a duplicate. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/22/07	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, sLastName, oRs, sResults

sLastName = DBsafe(request("userlname"))
sResults = ""

sSql = "SELECT userfname, userlname, useraddress FROM egov_users WHERE LOWER(userlname) = LOWER('" & sLastName
sSql = sSql & "') AND headofhousehold = 1 AND orgid = " & session("orgid") & " ORDER BY userlname, userfname, useraddress"

'response.write sSql & "<br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	Do While Not oRs.EOF
		sResults = sResults & vbcrlf & oRs("userfname") & " " & oRs("userlname") & " - " & oRs("useraddress")
		oRs.MoveNext
	Loop
Else
	sResults = "NEWCITIZEN"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults


'--------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	DBsafe = sNewString
End Function


%>
