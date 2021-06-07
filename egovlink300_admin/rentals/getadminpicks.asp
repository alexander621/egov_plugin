<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: getadminpicks.asp
' AUTHOR: Steve Loar
' CREATED: 10/26/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This gets the admin drop down using a name search. It is called via AJAX
'
' MODIFICATION HISTORY
' 1.0   10/26/2009   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearchName, sSql, oRs, sResults, sAdminIncl

sSearchName = dbsafe(request("searchname"))

If UserIsRootAdmin( Session("UserId") ) Then
	sAdminIncl = ""
Else
	sAdminIncl = " AND isrootadmin <> 1 "
End If 

sSql = "SELECT 1 AS foo, userid AS userid, firstname, lastname, ISNULL(email,'') AS email, "
sSql = sSql & " ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname "
sSql = sSql & " FROM users WHERE orgid = " & session("orgid")
sSql = sSql & " AND lastname LIKE '" & sSearchName & "%' " & sAdminIncl
sSql = sSql & " UNION "
sSql = sSql & " SELECT 2 AS foo, userid AS userid, firstname, lastname, ISNULL(email,'') AS email, "
sSql = sSql & " ISNULL(lastname,'') + ISNULL(firstname,'') AS sortname "
sSql = sSql & " FROM users WHERE orgid = " & session("orgid") 
sSql = sSql & " AND ( firstname LIKE '%" & sSearchName & "%' OR lastname LIKE '%" & sSearchName & "%' ) " & sAdminIncl
sSql = sSql & " AND userid NOT IN ( SELECT userid FROM users WHERE orgid = " & session("orgid")
sSql = sSql & " AND lastname LIKE '" & sSearchName & "%' )"
sSql = sSql & " ORDER BY foo, sortname, email, userid"
'response.write sSql

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If Not oRs.EOF Then
	sResults = "Select a Name: <select name='rentaluserid' id='rentaluserid'>"
	Do While Not oRs.EOF
		sResults = sResults & "<option value='" & oRs("userid") & "'"
		If iRentalUserid = CLng(oRs("userid")) Then
			sResults = sResults & " selected=""selected"" "
		End If 
		sResults = sResults & ">"
		sResults = sResults & oRs("lastname") & ", " & oRs("firstname")
		If oRs("email") <> "" Then
			sResults = sResults & " - " & oRs("email")
		End If 
		sResults = sResults & "</option>"
		oRs.MoveNext 
	Loop
	sResults = sResults & "</select>"
Else
	sResults = "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />No Matching Names Found"
End If 

oRs.Close
Set oRs = Nothing 

response.write sResults

%>