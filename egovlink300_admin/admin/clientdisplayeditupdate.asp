<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: clientdisplayeditupdate.asp
' AUTHOR: Steve Loar
' CREATED: 05/19/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Updates Displays. Called from displayedit.asp
'
' MODIFICATION HISTORY
' 1.0   05/19/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iDisplayId, sClientDisplayName, sClientDisplayDescription, sSql

If request("displayid") <> "" Then
	iDisplayId = CLng(request("displayid"))
Else
	response.redirect "displaylist.asp"
End If 

If request("usesdisplayname") = "0" Then
	sClientDisplayName = "NULL"
	sClientDisplayDescription = "'" & dbsafewithHTML(request("clientdisplaydescription")) & "'"
Else
	sClientDisplayName = "'" & dbsafewithHTML(request("clientdisplayname")) & "'"
	sClientDisplayDescription = "NULL"
End If 

' Remove the organization specific display
sSql = "DELETE FROM egov_organizations_to_displays WHERE displayid = " & iDisplayId
sSql = sSql & " AND orgid = " & session("orgid")
response.write sSql & "<br /><br />"
RunSQLStatement sSql

' Add the client display values
sSql = "INSERT INTO egov_organizations_to_displays ( displayid, orgid, displayname, displaydescription ) VALUES ( "
sSql = sSql & iDisplayId & ", " & session("orgid") & ", " & sClientDisplayName & ", " & sClientDisplayDescription & " )"
response.write sSql & "<br /><br />"

RunSQLStatement sSql

response.redirect "clientdisplayedit.asp?displayid=" & iDisplayId & "&msg=u" 

%>

