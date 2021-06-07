<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getpermitapplicants.asp
' AUTHOR: Steve Loar
' CREATED: 2/14/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the applicants drop down using a name search. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   2/14/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchName, sSql, oRs, sResults

sSearchName = dbsafe(request("searchname"))

'sSql = "SELECT userid, userfname, userlname, ISNULL(userbusinessname,'') AS userbusinessname, useraddress FROM egov_users "
'sSql = sSql & " WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 AND "
'sSql = sSql & " ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' OR userbusinessname LIKE '%" & sSearchName & "%' ) "
'sSql = sSql & " ORDER BY userlname, userfname, userbusinessname"

sSql = "SELECT userid AS userid, userbusinessname AS company, userfname AS firstname, userlname AS lastname, 'U' AS contacttype, "
sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') + ISNULL(userbusinessname,'') AS sortname, useraddress AS address "
sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' OR userbusinessname LIKE '%" & sSearchName & "%' ) "
sSql = sSql & " UNION "
sSql = sSql & " SELECT permitcontacttypeid AS userid, company AS company, firstname AS firstname, lastname AS lastname, 'C' AS contacttype, "
sSql = sSql & " ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname, address AS address "
sSql = sSql & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") 
sSql = sSql & " AND ( firstname LIKE '%" & sSearchName & "%' OR lastname LIKE '%" & sSearchName & "%' OR company LIKE '%" & sSearchName & "%' ) "
sSql = sSql & " ORDER BY sortname"

'response.write sSql & "<br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRS.EOF Then
	sResults = "<select name='userid' id='userid'>"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("contacttype") & oRs("userid") & "'>"
		If oRs("firstname") <> "" Then
			sResults = sResults & oRs("firstname") & " " & oRs("lastname")
			bName = True 
		Else
			bName = False 
		End If 

		If oRs("company") <> "" Then
			If bName Then 
				sResults = sResults &  " ("
			End If 
			sResults = sResults & oRs("company")
			If bName Then
				sResults = sResults & ")"
			End If 
		End If 

		If oRs("address") <> "" Then
			sResults = sResults & " - " & oRs("address")
		End If 

		sResults = sResults & "</option>"

		oRs.MoveNext
	Loop 
	sResults = sResults & "</select>"
Else
	sResults = "<input type='hidden' name='userid' id='userid' value='0' />No Match Found"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>
