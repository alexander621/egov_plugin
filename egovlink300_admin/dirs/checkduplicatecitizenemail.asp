<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkduplicatecitizenemail.asp
' AUTHOR: Steve Loar
' CREATED: 11/26/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This checks that the passed email is not a duplicate. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   11/26/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, sUserEmail, oRs, sResults

sUserEmail = DBsafe(request("email"))
sResults = ""

sSql = "SELECT COUNT(userid) AS hits FROM egov_users WHERE LOWER(useremail) = LOWER('" & sUserEmail
sSql = sSql & "') AND isdeleted = 0 AND headofhousehold = 1 AND orgid = " & session("orgid") 

'response.write sSql & "<br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If Not oRs.EOF Then 
	If CLng(oRs("hits")) > CLng(0) Then
		sResults = "DUPLICATE"
	Else
		sResults = "OK"
	End If 
Else
	sResults = "OK"
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