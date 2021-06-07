<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: agedaccountexport.asp
' AUTHOR: SteveLoar
' CREATED: 02/11/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  A report of citizen accounts with amounts over 1 year in their Memo Accounts
'
' MODIFICATION HISTORY
' 1.0   2/11/2013	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, toDate, fromDate

todaysDate = DateValue(Now)
agedDate = DateAdd("yyyy", -1, todaysDate)

sSql = "SELECT U.userid, ISNULL(U.userfname, '') AS userfname, ISNULL(U.userlname,'') AS userlname, ISNULL(U.useraddress,'') AS address, ISNULL(U.usercity,'') AS city, "
sSql = sSql & "ISNULL(U.userstate,'') AS state, ISNULL(U.userzip,'') AS zip, U.accountbalance, 0 AS deposits, U.accountbalance AS agedamount "
sSql = sSql & "FROM egov_users U "
sSql = sSql & "LEFT JOIN agedaccounthelper aa ON aa.accountid = u.userid and aa.paymentdate > '" & agedDate & "' and aa.orgid = u.orgid "
sSql = sSql & "WHERE U.orgid = " & session("orgid") & " AND U.accountbalance > 0 AND U.isdeleted = 0 and aa.accountid IS NULL "
'sSql = sSql & "AND U.userid NOT IN (SELECT A.accountid FROM egov_class_payment P, egov_accounts_ledger A WHERE P.orgid = " & session("orgid")
'sSql = sSql & " AND P.paymentdate > '" & agedDate & "' AND P.paymentid = A.paymentid AND A.paymenttypeid = 4 AND A.amount > 0) "
sSql = sSql & "UNION "
sSql = sSql & "SELECT U.userid, ISNULL(U.userfname, '') AS userfname, ISNULL(U.userlname,'') AS userlname, ISNULL(U.useraddress,'') AS address, ISNULL(U.usercity,'') AS city, "
sSql = sSql & "ISNULL(U.userstate,'') AS state, ISNULL(U.userzip,'') AS zip, U.accountbalance, SUM(A.amount) AS deposits, U.accountbalance - SUM(A.amount) AS agedamount "
sSql = sSql & "FROM egov_class_payment P, egov_accounts_ledger A, egov_users U "
sSql = sSql & "WHERE P.orgid = " & session("orgid") & " AND P.paymentdate > '" & agedDate & "' "
sSql = sSql & "AND P.paymentid = A.paymentid AND A.entrytype = 'credit' "
sSql = sSql & "AND A.paymenttypeid = 4 AND A.amount > 0 AND U.isdeleted = 0 AND P.orgid = U.orgid AND A.accountid = U.userid AND U.accountbalance > 0 "
sSql = sSql & "GROUP BY U.userid, U.userfname, U.userlname, U.useraddress, U.usercity, U.userstate, U.userzip, U.accountbalance "
sSql = sSql & "HAVING U.accountbalance > SUM(A.amount) "
sSql = sSql & "ORDER BY 3, 2"
'response.write sSql
'response.end


server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=AgedAccounts.xls"


response.write vbcrlf & "<html>"

response.write vbcrlf & "<style>  "
response.write " .moneystyle "
response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
response.write vbcrlf & "</style>"

response.write vbcrlf & "<body>"

'response.write sSql & "<br /><br />"

response.write vbcrlf & "<table border=""1"">"
response.write "<tr><td></td><td>Aged Account Report</td></tr>"
response.write "<tr><td></td><td>As Of " & todaysDate & "</td></tr>"
response.flush

' pull the report here
Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1



' if data found 
If Not oRs.EOF Then

	' print the title row
	response.write "<tr><th>User Id</th><th>First Name</th><th>Last Name</th><th>Address</th><th>City</th><th>State</th><th>Zip</th><th>Account Balance</th><th>Deposits Since " & agedDate & "</th><th>Aged Amount</th></tr>"
	response.flush

	Do While Not oRs.EOF
		response.write "<tr>"

		' UserId
		response.write "<td>" & oRs("userid") & "</td>"

		' First Name
		response.write "<td>" & oRs("userfname") & "</td>"
	
		' Last Name
		response.write "<td>" & oRs("userlname") & "</td>"

		' Address
		response.write "<td>" & oRs("address") & "</td>"

		' City
		response.write "<td>" & oRs("city") & "</td>"

		' State
		response.write "<td>" & oRs("state") & "</td>"

		' Zip
		response.write "<td>" & oRs("zip") & "</td>"

		' Account Balance
		response.write "<td align=""right"" class=""moneystyle"">" & oRs("accountbalance") & "</td>"

		' Deposits
		response.write "<td align=""right"" class=""moneystyle"">" & oRs("deposits") & "</td>"

		' Aged Amount
		response.write "<td align=""right"" class=""moneystyle"">" & oRs("agedamount") & "</td>"
		
		response.write "</tr>"
		response.flush

		oRs.MoveNext
	Loop 
Else 
	response.write "<tr><td></td><td>No accounts have deposits of more than 1 year at this time.</td></tr>"
	response.flush
End If 

oRs.Close
Set oRs = Nothing 



response.write vbcrlf & "</table>"
response.write vbcrlf & "</body></html>"
response.flush

%>


