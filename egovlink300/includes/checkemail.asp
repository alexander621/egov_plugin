<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: checkemail.asp
' AUTHOR: Steve Loar	
' CREATED: 01/29/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This check that the citizen email is not being used by another citizen.
'
' MODIFICATION HISTORY
' 1.0   01/29/2008	Steve Loar - Initial code 
' 1.1	01/14/2009	Steve Loar - Added check for not being deleted.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oEmail, sResponse

sSql = "SELECT COUNT(userid) AS hits FROM egov_users WHERE useremail = '" & Track_DBsafe(Trim(request("email")))
sSql = sSql & "' AND userid <> " & CLng(request("uid")) & " AND isdeleted = 0 AND orgid = " & CLng(request("orgid"))

Set oEmail = Server.CreateObject("ADODB.Recordset")
oEmail.Open sSQL, Application("DSN"), 3, 1

If NOT oEmail.EOF Then
	If clng(oEmail("hits")) = clng(0) Then 
		sResponse = "OK"
	Else
		sResponse = "TAKEN"
	End If 
Else
	sResponse = "OK"
End If 

oEmail.Close
Set oEmail = Nothing 

response.write sResponse


'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function Track_DBsafe( strDB )
'--------------------------------------------------------------------------------------------------
Function Track_DBsafe( strDB )
	Dim sNewString
	If Not VarType( strDB ) = vbString Then Track_DBsafe = strDB : Exit Function
	sNewString = Replace( strDB, "'", "''" )
	sNewString = Replace( sNewString, "<", "&lt;" )
	Track_DBsafe = sNewString
End Function


%>
