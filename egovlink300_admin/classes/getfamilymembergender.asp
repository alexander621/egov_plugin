<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getfamilymembergender.asp
' AUTHOR: Steve Loar
' CREATED: 10/11/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the gender of the passed family member, is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/11/2011	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, sGender

sSql = "SELECT ISNULL(U.gender,'N') AS gender "
sSql = sSql & "FROM egov_users U, egov_familymembers F "
sSql = sSql & "WHERE F.familymemberid = " & CLng(request("familymemberid")) & " AND F.userid = U.userid"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

If Not oRs.EOF Then
	sGender = oRs("gender")
Else
	sGender = "N"
End If 

oRs.Close
Set oRs = Nothing 

response.write sGender

%>