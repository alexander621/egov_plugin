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
Dim sSearchName, sSql, oRs, sResults, sOnChange

sSearchName = dbsafe(request("searchname"))
sSearchName2 = dbsafe(request("searchname2"))

If request("onchange") <> "none" Then
	sOnChange = " onchange='" & request("onchange") & "()' "
Else
	sOnChange = "" 
End If 

sSql = "SELECT 1 AS foo, userid, userfname AS firstname, userlname AS lastname, "
sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND (userlname LIKE '" & sSearchName & "%' OR userfname LIKE '" & sSearchName & "%') "
if sSearchName2 <> "" then sSql = sSql & " AND (userlname LIKE '" & sSearchName2 & "%' OR userfname LIKE '" & sSearchName2 & "%') "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 2 AS foo, userid, userfname AS firstname, userlname AS lastname, "
sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' ) "
if sSearchName2 <> "" then sSql = sSql & " AND ( userfname LIKE '%" & sSearchName2 & "%' OR userlname LIKE '%" & sSearchName2 & "%' ) "
sSql = sSql & " AND userid NOT IN ( SELECT userid FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 "
sSql = sSql & " AND headofhousehold = 1 AND userregistered = 1  "
sSql = sSql & " AND (userlname LIKE '" & sSearchName & "%' OR userfname LIKE '" & sSearchName & "%') "
if sSearchName2 <> "" then sSql = sSql & " AND (userlname LIKE '" & sSearchName2 & "%' OR userfname LIKE '" & sSearchName2 & "%') "
sSql = sSql & " ) "
sSql = sSql & " ORDER BY foo, sortname"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If Not oRS.EOF Then
	sResults = "Select a Name: <select name='egovuserid' id='egovuserid'" & sOnChange & ">"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("userid") & "'>"
		sResults = sResults & oRs("lastname") & ", " & oRs("firstname")
		If oRs("address") <> "" Then
			sResults = sResults & " - " & oRs("address")
		End If 

		sResults = sResults & "</option>"

		oRs.MoveNext
	Loop 
	sResults = sResults & "</select>"
Else
	sResults = "<input type='hidden' name='egovuserid' id='egovuserid' value='0' /><span class=""nomatch"">No Matching Names Found</span>"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>
