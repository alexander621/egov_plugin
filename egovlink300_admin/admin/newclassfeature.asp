<!-- #include file="../includes/common.asp" //-->
<%
' This script gives rights to the Classes and Events Reports for Orgs and Users
' Steve Loar, 4/6/2010

Dim sSql, oRs

sSql = "DELETE FROM egov_organizations_to_features WHERE featureid = 339"
response.write "<br /><br />" & sSql & "<br /><br />"
'RunSQLStatement sSql

sSql = "SELECT DISTINCT orgid FROM egov_organizations_to_features WHERE featureid IN (6,138)"
response.write "<hr /><br /><br />" & sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	sSql = "INSERT INTO egov_organizations_to_features ( featureid, orgid ) VALUES ( 339, " & oRs("orgid") & " )"
	response.write sSql & "<br />"
'	RunSQLStatement sSql

	oRs.MoveNext
Loop 

oRs.Close
Set oRs = Nothing 

sSql = "DELETE FROM egov_users_to_features WHERE featureid = 339"
response.write "<hr /><br /><br />" & sSql & "<br /><br />"
'RunSQLStatement sSql


sSql = "SELECT DISTINCT userid FROM egov_users_to_features WHERE featureid IN (6,138)"
response.write "<hr /><br /><br />" & sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

Do While Not oRs.EOF
	sSql = "INSERT INTO egov_users_to_features ( featureid, permissionid, userid ) VALUES ( 339, 1, " & oRs("userid") & " )"
	response.write sSql & "<br />"
'	RunSQLStatement sSql

	oRs.MoveNext
Loop 

oRs.Close
Set oRs = Nothing 

response.write "<hr /><br /><br />Finished"

%>