<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getcitizenpicks.asp
' AUTHOR: Steve Loar
' CREATED: 9/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the citizens drop down using a name search. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   09/28/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchName, sSql, oRs, sResults

sSearchName = dbsafe(request("searchname"))


sSql = "SELECT 1 AS foo, userid, userfname AS firstname, userlname AS lastname, "
sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND userlname LIKE '" & sSearchName & "%' "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 2 AS foo, userid, userfname AS firstname, userlname AS lastname, "
sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' ) "
sSql = sSql & " AND userid NOT IN ( SELECT userid FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 "
sSql = sSql & " AND headofhousehold = 1 AND userregistered = 1 AND userlname LIKE '" & sSearchName & "%' ) "
sSql = sSql & " ORDER BY foo, sortname"

sSql = "SELECT 1 AS foo, userid, userfname AS firstname, userlname AS lastname, "
sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address, "
sSql = sSql & " residenttypename = CASE WHEN R.description IS NULL THEN '' ELSE R.description END "
sSql = sSql & " FROM egov_users U LEFT OUTER JOIN egov_poolpassresidenttypes R ON U.residenttype = R.resident_type AND U.orgid = R.orgid "
sSql = sSql & " WHERE U.orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND userlname LIKE '" & sSearchName & "%' "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 2 AS foo, userid, userfname AS firstname, userlname AS lastname, "
sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address, "
sSql = sSql & " residenttypename = CASE WHEN R2.description IS NULL THEN '' ELSE R2.description END "
sSql = sSql & " FROM egov_users U2 LEFT OUTER JOIN egov_poolpassresidenttypes R2 ON U2.residenttype = R2.resident_type AND U2.orgid = R2.orgid "
sSql = sSql & " WHERE U2.orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' ) "
sSql = sSql & " AND userid NOT IN ( SELECT userid FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 "
sSql = sSql & " AND headofhousehold = 1 AND userregistered = 1 AND userlname LIKE '" & sSearchName & "%' ) "
sSql = sSql & " ORDER BY foo, sortname"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 0, 1

If Not oRS.EOF Then
	sResults = "<label for='userid'>Select a Name:</label> <select name='userid' id='userid' onchange='setPricePicks();'>"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("userid") & "'>"
		sResults = sResults & oRs("lastname") & ", " & oRs("firstname")
		If oRs("residenttypename") <> "" Then
			sResults = sResults & " (" & oRs("residenttypename") & ")"
		End If 
		If oRs("address") <> "" Then
			sResults = sResults & " - " & oRs("address")
		End If 

		sResults = sResults & "</option>"

		oRs.MoveNext
	Loop 
	sResults = sResults & "</select>"
Else
	sResults = "<input type='hidden' name='userid' id='userid' value='0' /><span id=""nomatch"">No Matching Names Found</span>"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>