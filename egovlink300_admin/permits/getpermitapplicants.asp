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
' 1.0   02/14/2008   Steve Loar - INITIAL VERSION
' 2.0	09/23/2008	Steve Loar - Changed to pull contact types as well as registered users
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchName, sSql, oRs, sResults

strFieldName = "userid"
if request.querystring("fieldname") <> "" then strFieldName = request.querystring("fieldname")

sSearchName = dbsafe(request("searchname"))
sSearchCompanyName = dbsafe(request("searchcompanyname"))
if sSearchCompanyName = "" then sSearchCompanyName = "234234#@#$@#$asdfadsf"
sSearchFirstName = dbsafe(request("searchfirstname"))
if sSearchFirstName = "" then sSearchFirstName = "234234#@#$@#$asdfadsf"
sSearchLastName = dbsafe(request("searchlastname"))
if sSearchLastName = "" then sSearchLastName = "234234#@#$@#$asdfadsf"

'sSql = "SELECT userid AS userid, userbusinessname AS company, userfname AS firstname, userlname AS lastname, 'U' AS contacttype, "
'sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') + ISNULL(userbusinessname,'') AS sortname, useraddress AS address "
'sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
'sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' OR userbusinessname LIKE '%" & sSearchName & "%' ) "
'sSql = sSql & " UNION "
'sSql = sSql & " SELECT permitcontacttypeid AS userid, company AS company, firstname AS firstname, lastname AS lastname, 'C' AS contacttype, "
'sSql = sSql & " ISNULL(lastname,'') + ISNULL(firstname,'') + ISNULL(company,'') AS sortname, address AS address "
'sSql = sSql & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") 
'sSql = sSql & " AND ( firstname LIKE '%" & sSearchName & "%' OR lastname LIKE '%" & sSearchName & "%' OR company LIKE '%" & sSearchName & "%' ) "
'sSql = sSql & " ORDER BY sortname"
sSql = ""
if request.querystring("contactsonly") <> "yes" then
sSql = sSql & "SELECT 1 AS foo, userid AS userid, userbusinessname AS company, userfname AS firstname, userlname AS lastname, 'U' AS contacttype, "
sSql = sSql & " ISNULL(userbusinessname,'') + ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND userlname LIKE '" & sSearchName & "%' "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 2 AS foo, userid AS userid, userbusinessname AS company, userfname AS firstname, userlname AS lastname, 'U' AS contacttype, "
sSql = sSql & " ISNULL(userbusinessname,'') + ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' OR userbusinessname LIKE '%" & sSearchName & "%' ) "
sSql = sSql & " AND userid NOT IN (SELECT userid FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
sSql = sSql & " AND userlname LIKE '" & sSearchName & "%' ) "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 1 AS foo, permitcontacttypeid AS userid, company AS company, firstname AS firstname, lastname AS lastname, 'C' AS contacttype, "
sSql = sSql & " ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, address AS address "
sSql = sSql & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") 
sSql = sSql & " AND lastname IS NULL AND company LIKE '" & sSearchName & "%' "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 2 AS foo, permitcontacttypeid AS userid, company AS company, firstname AS firstname, lastname AS lastname, 'C' AS contacttype, "
sSql = sSql & " ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, address AS address "
sSql = sSql & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") 
sSql = sSql & " AND ( firstname LIKE '%" & sSearchName & "%' OR lastname LIKE '%" & sSearchName & "%' OR company LIKE '%" & sSearchName & "%' ) AND "
sSql = sSql & " permitcontacttypeid NOT IN (SELECT permitcontacttypeid FROM egov_permitcontacttypes WHERE orgid = " & session("orgid")
sSql = sSql & " AND lastname IS NULL AND company LIKE '" & sSearchName & "%' ) "
sSql = sSql & " ORDER BY foo, sortname"
else
sSql = sSql & " SELECT 1 AS foo, permitcontacttypeid AS userid, company AS company, firstname AS firstname, lastname AS lastname, 'C' AS contacttype, "
sSql = sSql & " ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, address AS address "
sSql = sSql & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") 
sSql = sSql & " AND lastname IS NULL AND company LIKE '" & sSearchCompanyName & "%' "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 2 AS foo, permitcontacttypeid AS userid, company AS company, firstname AS firstname, lastname AS lastname, 'C' AS contacttype, "
sSql = sSql & " ISNULL(company,'') + ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname, address AS address "
sSql = sSql & " FROM egov_permitcontacttypes WHERE orgid = " & session("orgid") 
sSql = sSql & " AND ( firstname LIKE '%" & sSearchFirstName & "%' OR lastname LIKE '%" & sSearchLastName & "%' OR company LIKE '%" & sSearchCompanyName & "%' ) AND "
sSql = sSql & " permitcontacttypeid NOT IN (SELECT permitcontacttypeid FROM egov_permitcontacttypes WHERE orgid = " & session("orgid")
sSql = sSql & " AND lastname IS NULL AND company LIKE '" & sSearchCompanyName & "%' ) "
sSql = sSql & " ORDER BY foo, sortname"
end if



'response.write sSql & "<br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRS.EOF Then
	sResults = "<select name='" & strFieldName & "' id='" & strFieldName & "' onchange=""toggleAddressSearch();"">"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("contacttype") & oRs("userid") & "'>"

		If oRs("company") <> "" Then
			sResults = sResults &  oRs("company")
			If oRs("firstname") <> "" Then 
				sResults = sResults &  " - "
			End If 
		End If 
		If oRs("firstname") <> "" Then 
			sResults = sResults &  oRs("lastname") & ", " & oRs("firstname")
		End If 

'		If oRs("firstname") <> "" Then
'			sResults = sResults & oRs("lastname") & ", " & oRs("firstname")
'			bName = True 
'		Else
'			bName = False 
'		End If 
'
'		If oRs("company") <> "" Then
'			If bName Then 
'				sResults = sResults &  " ("
'			End If 
'			sResults = sResults & oRs("company")
'			If bName Then
'				sResults = sResults & ")"
'			End If 
'		End If 

		If oRs("address") <> "" Then
			sResults = sResults & " - " & oRs("address")
		End If 

		sResults = sResults & "</option>"

		oRs.MoveNext
	Loop 
	sResults = sResults & "</select>"
Else
	sResults = "<input type='hidden' name='" & strFieldName & "' id='" & strFieldName & "' value='0' />No Match Found"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>
