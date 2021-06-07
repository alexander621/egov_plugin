<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: regattateamexport.asp
' AUTHOR: Steve Loar
' CREATED: 04/08/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Extract of selected regatta teams exported to EXCEL
'
' MODIFICATION HISTORY
' 1.0   04/08/2009	Steve Loar - INITIAL VERSION
' 1.1	5/14/2010	Steve Loar - Split captain name into first and last, AND changed order of export fields
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTeamPicks, sRptTitle, sSql, oRs

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())

server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=River_Regatta_Teamss.xls"

sTeamPicks = request("teampicks")

sRptTitle = vbcrlf & "<tr height=""30""><th>River Regatta Teams</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

'sSql = "SELECT T.regattateamid, T.regattateam, T.captainname, ISNULL(G.regattateamgroup,'') AS regattateamgroup, "
'sSql = sSql & " M.regattateammember, M.classlistid, L.paymentid, ISNULL(L.quantity,1) AS quantity, P.paymentdate, P.userid, AL.amount, L.status, "
'sSql = sSql & " U.userfname, U.userlname, U.useraddress, U.usercity, U.userstate, U.userzip, ISNULL(U.userhomephone,'') AS userhomephone, "
'sSql = sSql & " ISNULL(U.emergencycontact,'') AS emergencycontact, ISNULL(U.emergencyphone,'') AS emergencyphone, U.useremail "
'sSql = sSql & " FROM egov_regattateams T LEFT OUTER JOIN egov_regattateamgroups G "
'sSql = sSql & " ON T.orgid = G.orgid AND T.membercount >= G.minteamcount AND T.membercount <= G.maxteamcount "
'sSql = sSql & " INNER JOIN egov_regattateammembers M ON T.orgid = M.orgid AND T.regattateamid = M.regattateamid "
'sSql = sSql & " INNER JOIN egov_class_list L ON M.classlistid = L.classlistid "
'sSql = sSql & " INNER JOIN egov_class_payment P ON L.paymentid = P.paymentid "
'sSql = sSql & " INNER JOIN egov_accounts_ledger AL ON L.paymentid = AL.paymentid AND L.classlistid = AL.itemid "
'sSql = sSql & " INNER JOIN egov_users U ON P.userid = U.userid "
'sSql = sSql & " WHERE T.orgid = " & session("orgid") & " AND T.regattateamid IN " & sTeamPicks
'sSql = sSql & " ORDER BY T.regattateam, T.regattateamid, M.regattateammember"

sSql = "SELECT T.regattateamid, T.regattateam, G.regattateamgroup, T.captainfirstname, T.captainlastname, T.captainaddress, T.captaincity, "
sSql = sSql & "T.captainstate, T.captainzip, T.captainphone,U.userfname, U.userlname, U.useremail "
sSql = sSql & "FROM egov_regattateams T, egov_regattateamgroups G, egov_class_list C, egov_class_payment P, egov_users U "
sSql = sSql & "WHERE T.regattateamgroupid = G.regattateamgroupid AND T.classlistid = C.classlistid "
sSql = sSql & "AND C.paymentid = P.paymentid AND P.userid = U.userid "
sSql = sSql & "AND T.orgid = " & session("orgid") & " AND T.regattateamid IN " & sTeamPicks
sSql = sSql & "ORDER BY T.regattateam, G.regattateamgroup, T.regattateamid"
'response.write sSql & "<br /><br />"


Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 0, 1

If Not oRs.EOF Then 
	response.write "<html><body>"
	response.write "<table cellpadding=""4"" cellspacing=""0"" border=""1"">"
	'response.write "<tr height=""30""><th>River Regatta Teams</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	response.write "<tr height=""30""><th>Captain First Name</th><th>Captain Last Name</th><th>Captain Phone</th>"
	response.write "<th>Captain Address</th><th>Captain City</th><th>Captain State</th><th>Captain Zip</th>"
	response.write "<th>Email</th><th>Payee Name</th><th>Team Name</th><th>Team Category</th>"
	response.write "</tr>"
	response.flush
	
	Do While Not oRs.EOF
		response.write "<tr height=""20"">"
		response.write "<td align=""center"">" & oRs("captainfirstname") & "</td>"
		response.write "<td align=""center"">" & oRs("captainlastname") & "</td>"
		response.write "<td align=""center"">&nbsp;" & formatphonenumber(oRs("captainphone")) & "</td>"
		response.write "<td align=""left"">" & oRs("captainaddress") & "</td>"
		response.write "<td align=""center"">" & oRs("captaincity") & "</td>"
		response.write "<td align=""center"">" & oRs("captainstate") & "</td>"
		response.write "<td align=""center"">&nbsp;" & oRs("captainzip") & "</td>"
		response.write "<td align=""center"">" & oRs("useremail") & "</td>"
		response.write "<td align=""center"">" & Trim(oRs("userfname") & " " & oRs("userlname")) & "</td>"
		response.write "<td align=""center"">" & oRs("regattateam") & "</td>"
		response.write "<td align=""center"">" & oRs("regattateamgroup") & "</td>"
		response.write "</tr>"
		response.flush

		oRs.MoveNext
	Loop 

	response.write "</table></body></html>"
	response.flush
End If 

oRs.Close
Set oRs = Nothing 

%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
