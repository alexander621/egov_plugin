<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattateamnameupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/05/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates a Regatta team name. Called via AJAX from regattateamnameedit.asp
'
' MODIFICATION HISTORY
' 1.0   08/05/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRegattaTeamId, sTeamName, sSql

If request("regattateamid") = "" Or request("regattateam") = "" Then
	response.write "Usage Error"
	response.End 
End If 

iRegattaTeamId = CLng(request("regattateamid"))
sTeamName = request("regattateam")

sSql = "UPDATE egov_regattateams SET regattateam = '" & dbsafe(sTeamName) & "'"
sSql = sSql & " WHERE regattateamid = " & iRegattaTeamId & " AND orgid = " & session("orgid")
'response.write sSql & "<br />"
RunSQLStatement sSql

'response.write "Success"


%>