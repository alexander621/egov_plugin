<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: account_distribution_export.asp
' AUTHOR: Steve Loar
' CREATED: 12/18/2009
' COPYRIGHT: Copyright 2090 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   12/18/2009	Steve Loar - INITIAL VERSION
' 1.1	05/21/2010	Steve Loar - Changed to style numbers as currency in excel
' 1.2	08/27/2010	Steve Loar - Added Citizen Account activity that was left out
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, oRs, oSchema, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal
Dim iLocationId, toDate, fromDate, sDateRange, iPaymentLocationId, iReportType, sAdminlocation
Dim sFile, sRptTitle, sWhereClause, iJournalEntryTypeId, today, sNameClause, iReservationTypeId
Dim sCitizenNameClause, iAccountNo
Dim from_time, to_time, where_time

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
sWhereClause = ""

If request("reporttype") = "" Then 
	' Summary Report is the default
	iReportType = CLng(1)
Else
	iReportType = CLng(request("reporttype"))
End If 

If iReportType = CLng(1) Then
	sRptTitle = "<tr><th></th><th>Account Distribution Summary - Rentals</th><th></th><th></th><th></th></tr>"
	sFile = "Summary_"
	sRptType = "Summary"
Else
	If iReportType = CLng(2) Then
		sRptTitle = "<tr><th></th><th>Account Distribution Detail - Rentals</th><th></th><th></th><th></th></tr>"
		sFile = "Detail_"
		sRptType = "Detail"
	Else
		sRptTitle = "<tr><th></th><th>Account Distribution List - Rentals</th><th></th><th></th><th></th></tr>"
		sFile = "List_"
		sRptType = "List"
	End If 
End If 

' These next 4 lines generate the Excel spreadsheet
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Account_Distribution_" & sFile & sDate & ".xls"

' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then
	toDate = today 
End If

If fromDate = "" or IsNull(fromDate) Then 
	fromDate = today
End If

If Request("fromtime") <> "" Then 
	from_time = Request("fromtime")
Else 
	from_time = "none"
End If 
If Request("totime") <> "" Then 
	to_time = Request("totime")
Else
	to_time = "none"
End If 

If request("locationid") = "0" Then
	iLocationId = 0
Else
	iLocationId = CLng(request("locationid"))
End If 

If request("paymentlocationid") = "" Then
	iPaymentLocationId = 0
Else
	iPaymentLocationId = CLng(request("paymentlocationid"))
End If 

If request("adminuserid") = "" Then
	iAdminUserId = 0
Else
	iAdminUserId = CLng(request("adminuserid"))
End If 

If request("reservationtypeid") <> "" Then
	iReservationTypeId = CLng(request("reservationtypeid"))
Else
	iReservationTypeId = CLng(2)
End If 

If request("accountid") = "" Then
	iAccountNo = 0
Else
	iAccountNo = CLng(request("accountid"))
End If 


' BUILD SQL WHERE CLAUSE

sWhereClause = sWhereClause & " AND P.orgid = " & session("orgid") 

'sWhereClause = " AND (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
'sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & fromDate & "</th><th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th></tr>"
If from_time = "none" Then 
	sWhereClause = sWhereClause & " AND paymentDate >= '" & fromDate & "' "
	sRptTitle = sRptTitle & "<tr><th>Transaction Date >= " & fromDate & "</th>"
Else
	where_time = CDate( fromdate & " " & from_time )
	sWhereClause = sWhereClause & " AND paymentDate >= '" & where_time & "' "
	sRptTitle = sRptTitle & "<tr><th>Transaction Date >= " & where_time & "</th>"
End If 

If to_time = "none" Then 
	sWhereClause = sWhereClause & " AND paymentDate <= '" & DateAdd("d",1,toDate) & "' "
	sRptTitle = sRptTitle & "<th>AND Transaction Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th></tr>"
Else 
	where_time = CDate( todate & " " & to_time )
	sWhereClause = sWhereClause & " AND paymentDate <= '" & where_time & "' "
	sRptTitle = sRptTitle & "<th>AND Transaction Date <= " & where_time & "</th><th></th><th></th><th></th></tr>"
End If 

If iLocationId > 0 Then
	sWhereClause = sWhereClause & " AND adminlocationid = " & iLocationId
	sRptTitle = sRptTitle & "<tr><th>Admin Location: " & GetLocationName( iLocationId )  & "</th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Admin Location: All Locations</th><th></th><th></th><th></th><th></th></tr>"
End If 

If CLng(iAdminUserId) > CLng(0) Then
	sWhereClause = sWhereClause & " AND P.adminuserid = " & iAdminUserId
	sRptTitle = sRptTitle & "<tr><th>Admin: " & GetAdminName( iAdminUserId )  & "</th><th></th><th></th><th></th><th></th></tr>"
Else 
	sRptTitle = sRptTitle & "<tr><th>Admin: All Admins</th><th></th><th></th><th></th><th></th></tr>"
End If 

If iPaymentLocationId > 0 Then
	If iPaymentLocationId = CLng(2) Then
		sWhereClause = sWhereClause & " AND P.paymentlocationid = 3 " 
		sRptTitle = sRptTitle & "<tr><th>Payment Location: Web Site</th><th></th><th></th><th></th><th></th></tr>"
	Else
		sWhereClause = sWhereClause & " AND P.paymentlocationid < 3 " 
		sRptTitle = sRptTitle & "<tr><th>Payment Location: Office</th><th></th><th></th><th></th><th></th></tr>"
	End If 
Else
	sRptTitle = sRptTitle & "<tr><th>Payment Location: All Locations</th><th></th><th></th><th></th><th></th></tr>"
End If 

If request("journalentrytypeid") = "" Then
	iJournalEntryTypeId = 0
Else
	iJournalEntryTypeId = CLng(request("journalentrytypeid"))
End If 

If request("namelike") <> "" Then
	sNameClause = " AND (EU.userlname LIKE '%" & dbsafe( request("namelike") ) & "%' OR U.LastName LIKE '%" & dbsafe( request("namelike") ) & "%') "
	sCitizenNameClause = " AND EU.userlname LIKE '%" & dbsafe( request("namelike") ) & "%' "
Else
	sNameClause = ""
	sCitizenNameClause = ""
End If 

If iJournalEntryTypeId > 0 Then 
	sWhereClause = sWhereClause & " AND P.journalentrytypeid = " & iJournalEntryTypeId
	sRptTitle = sRptTitle & "<tr><th>Entries: " & GetJournalEntryDisplay( iJournalEntryTypeId ) & "</th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Entries: Payments and Refunds</th><th></th><th></th><th></th><th></th></tr>"
End If 

If iReservationTypeId > CLng(0) Then 
	sWhereClause = sWhereClause & " AND R.reservationtypeid = " & iReservationTypeId & " "
	sRptTitle = sRptTitle & "<tr><th>Reservation Type: " & GetReservationType( iReservationTypeId ) & " Only</th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Reservation Type: All Reservation Types</th><th></th><th></th><th></th><th></th></tr>"
End If 

If OrgHasFeature("gl accounts") Then 
	If CLng(iAccountNo) > CLng(0) Then
		sWhereClause = sWhereClause & " AND A.accountid = " & iAccountNo & " "
		sRptTitle = sRptTitle & "<tr><th>GL Account: " & GetAccountName( iAccountNo ) & " Only</th><th></th><th></th><th></th></tr>"
	Else
		sRptTitle = sRptTitle & "<tr><th>GL Account: All GL Accounts</th><th></th><th></th><th></th><th></th></tr>"
	End If 
End If 

If LCase(sRptType) = "summary" Then
	DisplaySummary sWhereClause, sRptTitle
Else
	DisplayDetails sWhereClause, sRptTitle, sNameClause, LCase(sRptType), sCitizenNameClause
End If 


'--------------------------------------------------------------------------------------------------
' void DisplaySummary varWhereClause, sRptTitle 
'--------------------------------------------------------------------------------------------------
Sub DisplaySummary( ByVal sWhereClause, ByVal sRptTitle )
	Dim bHasData, sSql, oRs

	iOldAccountId = CLng(0) 
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)
	bHasData = False 

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	'oSchema.fields.append "accountid", adInteger, , adFldUpdatable
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "creditamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "debitamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "totalamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oSchema.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount, 0 AS iscitizenaccount, "
	sSql = sSql & "SUM(L.amount) AS amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_rentalreservations R "
	sSql = sSql & " WHERE A.accountid = L.accountid AND L.paymentid = P.paymentid "
	sSql = sSql & " AND L.amount >= 0.00 AND P.isforrentals = 1 AND P.reservationid = R.reservationid AND A.orgid = P.orgid " & sWhereClause
	sSql = sSql & " GROUP BY A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount "
	sSql = sSql & " ORDER BY A.accountid, L.entrytype"
	'response.write "<tr><td>" & sSql & "<br /><br /></td></tr>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) = iOldAccountId Then
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = oSchema("creditamt") + oRs("amount")
					dTotal = dTotal + CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					oSchema("totalamt") = dTotal 
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRs("amount"))
					dTotal = dTotal - CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal - CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					oSchema("totalamt") = dTotal 
				End If 
			Else
				oSchema.addnew 
				oSchema("accountname") = oRs("accountname")
				oSchema("accountnumber") = oRs("accountnumber")
				oSchema("ispaymentaccount") = oRs("ispaymentaccount")
				oSchema("iscitizenaccount") = False 
				oSchema("creditamt") = 0.00
				oSchema("debitamt") = 0.00
				oSchema("totalamt") = 0.00
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					dTotal = CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					oSchema("totalamt") = oRs("amount")
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					dTotal = -CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal  - CDbl(oRs("amount"))
					oSchema("totalamt") = -CDbl(oRs("amount"))
				End If 
				iOldAccountId = CLng(oRs("accountid"))
			End If 
			oSchema.Update
			oRs.MoveNext
		Loop
	End If 
	oRs.Close
	Set oRs = Nothing


	'Get the citizen accounts summary here
	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount, 1 AS iscitizenaccount, SUM(L.amount) AS amount "
	sSql = sSql & "FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, "
	sSql = sSql & "egov_organizations_to_paymenttypes OP, egov_rentalreservations R "
	sSql = sSql & "WHERE L.paymentid = P.paymentid AND L.paymenttypeid = 4 AND L.amount >= 0.00 AND "
	sSql = sSql & "P.isforrentals = 1 AND P.reservationid = R.reservationid "
	sSql = sSql & "AND A.accountid = OP.accountid AND OP.paymenttypeid = L.paymenttypeid AND A.orgid = P.orgid " & sWhereClause
	sSql = sSql & " GROUP BY A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount "
	sSql = sSql & "ORDER BY A.accountid, L.entrytype"
	'response.write "<tr><td>" & sSql & "<br /><br /></td></tr>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF then
		bHasData = True 
		iOldAccountId = CLng(0)

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) = iOldAccountId Then
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = oSchema("creditamt") + oRs("amount")
					dTotal = dTotal + CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					oSchema("totalamt") = dTotal 
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRs("amount"))
					dTotal = dTotal - CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal - CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					oSchema("totalamt") = dTotal 
				End If 
			Else 
				oSchema.addnew 
				oSchema("accountname") = oRs("accountname") 
				oSchema("accountnumber") = oRs("accountnumber")
				oSchema("ispaymentaccount") = True 
				oSchema("iscitizenaccount") = True 
				oSchema("creditamt") = 0.00
				oSchema("debitamt") = 0.00
				oSchema("totalamt") = 0.00
				iOldAccountId = CLng(oRs("accountid"))
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					dTotal = CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					oSchema("totalamt") = oRs("amount")
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					dTotal = -CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal  - CDbl(oRs("amount"))
					oSchema("totalamt") = oRs("amount")
				End If 
			End If 
				
			oSchema.Update
			oRs.MoveNext
		Loop
	End If 

	oRs.Close
	Set oRs = Nothing 


	If Not bHasData Then 
		' A blank row
		oSchema.addnew 
		'oSchema("accountid") = 0
		oSchema("accountname") = " "
		oSchema("accountnumber") = " "
		oSchema("creditamt") = 0.00
		oSchema("debitamt") = 0.00
		oSchema("totalamt") = 0.00
		oSchema.Update
	End If 

	' Sort them 
	oSchema.Sort = "ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC"

	' Total Row
	sTotalRow = "<tr><td></td><td>Total</td><td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalCredit, 2,,,0) & "</td><td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalDebit, 2,,,0) & "</td><td align=""right"" class=""moneystyle"">" & FormatNumber(dGrandTotal,2,,,0) & "</td></tr>"

	oSchema.MoveFirst

	'CreateExcelDownload sRptTitle, sTotalRow

	If Not oSchema.EOF Then
		response.write "<html>"
		
		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"

		response.write "<body><table border=""0"">"

		'response.write "<tr><th></th><th>Account Distribution Summary</th><th></th><th></th><th></th></tr>"

		response.write sRptTitle
		response.flush

		response.write "<tr>"
		' WRITE COLUMN HEADINGS
		response.write  "<th>Account Name</th>"
		response.write  "<th>Account Number</th>"
		response.write  "<th>Total Amt Credited</th>"
		response.write  "<th>Total Amt Debited</th>"
		response.write  "<th>Total Amt Transfered</th>"
		response.write "</tr>"

		Do While Not oSchema.EOF
			response.write "<tr>"

			' Account Name
			response.write "<td align=""left"">&nbsp;" & oSchema("accountname") & "</td>"

			' Account Number
			response.write "<td align=""center"">&nbsp;" & oSchema("accountnumber") & "</td>"

			' Total Amt Credited
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("creditamt"),2,,,0) & "</td>"

			' Total Amt Debited
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("debitamt"),2,,,0) & "</td>"

			' Total Amt Transfered
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("totalamt"),2,,,0) & "</td>"

			response.write "</tr>"
			response.flush
			oSchema.MoveNext
		Loop

	End If 

	response.write sTotalRow
	response.flush

	response.write "</table></body></html>"
	response.flush

	oSchema.Close
	Set oSchema = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayDetails sWhereClause, sRptTitle, sNameClause, sRptType, sCitizenNameClause
'--------------------------------------------------------------------------------------------------
Sub DisplayDetails( ByVal sWhereClause, ByVal sRptTitle, ByVal sNameClause, ByVal sRptType, ByVal sCitizenNameClause )
	Dim bHasData, sSql, oRs

	iOldAccountId = CLng(0) 
	iOldPaymentId = CLng(0)
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)
	bHasData = False 

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	oSchema.fields.append "accountid", adInteger, , adFldUpdatable
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "receiptno", adInteger, , adFldUpdatable
	oSchema.fields.append "paymentdate", adDBTimeStamp, , adFldUpdatable
	oSchema.fields.append "paymenttime", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "creditamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "debitamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "totalamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oSchema.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable
	oSchema.fields.append "rentername", adVarChar, 100, adFldUpdatable
	oSchema.fields.append "reservationdate", adDBTimeStamp, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate, P.reservationid, "
	sSql = sSql & " T.reservationtypeselector, ISNULL(L.paymenttypeid,0) AS paymenttypeid, P.userid, "
	sSql = sSql & " P.journalentrytypeid, L.ispaymentaccount, 0 AS iscitizenaccount, "
	sSql = sSql & " CASE T.reservationtypeselector WHEN 'public' THEN EU.userfname ELSE U.FirstName END AS renterfirstname, "
	sSql = sSql & " CASE T.reservationtypeselector WHEN 'public' THEN EU.userlname ELSE U.LastName END AS renterlastname "
	sSql = sSql & " FROM egov_class_payment P INNER JOIN "
	sSql = sSql & " egov_accounts_ledger L ON P.paymentid = L.paymentid INNER JOIN "
	sSql = sSql & " egov_accounts A ON L.accountid = A.accountid INNER JOIN "
	sSql = sSql & " egov_rentalreservations R ON R.reservationid = P.reservationid INNER JOIN "
	sSql = sSql & " egov_rentalreservationtypes T ON R.reservationtypeid = T.reservationtypeid LEFT OUTER JOIN "
	sSql = sSql & " egov_users EU ON P.userid = EU.userid LEFT OUTER JOIN "
	sSql = sSql & " Users U ON P.userid = U.UserID "
	sSql = sSql & " WHERE L.amount >= 0.00 AND P.isforrentals = 1 " & sWhereClause & sNameClause
	sSql = sSql & " ORDER BY A.accountid, P.paymentid, L.entrytype"
	'response.write "<tr><td>" & sSql & "<br /><br /></td></tr>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) = iOldAccountId And CLng(oRs("paymentid")) = iOldPaymentId Then
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRs("amount"))
					dTotal = dTotal + CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					oSchema("totalamt") = oSchema("totalamt") + CDbl(oRs("amount")) 
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRs("amount"))
					dTotal = dTotal - CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal - CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					oSchema("totalamt") = oSchema("totalamt") - CDbl(oRs("amount")) 
				End If 
			Else
				oSchema.addnew 
				oSchema("accountid") = oRs("accountid")
				oSchema("accountname") = oRs("accountname")
				oSchema("accountnumber") = oRs("accountnumber")
				oSchema("ispaymentaccount") = oRs("ispaymentaccount")
				If oRs("accountname") = "Citizen Accounts" Then
					oSchema("iscitizenaccount") = True 
				Else 
					oSchema("iscitizenaccount") = False 
				End If 
				oSchema("receiptno") = oRs("paymentid")
				oSchema("paymentdate") = FormatDateTime(oRs("paymentdate"),2)
				oSchema("paymenttime") = FormatDateTime(oRs("paymentdate"),3)
				oSchema("creditamt") = 0.00
				oSchema("debitamt") = 0.00
				oSchema("totalamt") = 0.00
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					dTotal = CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					oSchema("totalamt") = oRs("amount")
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = -CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					dTotal = - CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal  - CDbl(oRs("amount"))
					oSchema("totalamt") = - CDbl(oRs("amount"))
				End If 
				iOldAccountId = CLng(oRs("accountid"))
				iOldPaymentId = CLng(oRs("paymentid"))
				oSchema("rentername") = Trim(oRs("renterfirstname") & " " & oRs("renterlastname"))
				oSchema("reservationdate") = DateValue(GetLastReservationDate( oRs("reservationid") ))
			End If 
			oSchema.Update
			oRs.MoveNext
		Loop
	End If 
	oRs.Close
	Set oRs = Nothing

	'Get the citizen accounts details here
	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate, "
	sSql = sSql & " P.reservationid, T.reservationtypeselector, ISNULL(L.paymenttypeid,0) AS paymenttypeid, P.userid, "
	sSql = sSql & " P.journalentrytypeid, L.ispaymentaccount, 1 AS iscitizenaccount, EU.userfname AS renterfirstname, "
	sSql = sSql & " EU.userlname AS renterlastname "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment P, egov_accounts A, egov_organizations_to_paymenttypes OP, "
	sSql = sSql & " egov_rentalreservations R, egov_rentalreservationtypes T, egov_users EU "
	sSql = sSql & " WHERE L.paymentid = P.paymentid AND L.paymenttypeid = 4 AND L.amount >= 0.00 AND P.isforrentals = 1 "
	sSql = sSql & " AND R.reservationtypeid = T.reservationtypeid AND P.userid = EU.userid AND A.accountid = OP.accountid "
	sSql = sSql & " AND OP.paymenttypeid = L.paymenttypeid AND OP.orgid = P.orgid AND P.reservationid = R.reservationid "
	sSql = sSql & sWhereClause & sCitizenNameClause
	sSql = sSql & " ORDER BY A.accountid, P.paymentid, L.entrytype"	
	'response.write "<tr><td>" & sSql & "</td></tr>"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		bHasData = True 
		iOldAccountId = CLng(0)
		iOldPaymentId = CLng(0)

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) = iOldAccountId And CLng(oRs("paymentid")) = iOldPaymentId Then
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRs("amount"))
					dTotal = dTotal + CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
				oSchema("totalamt") = oSchema("totalamt") + CDbl(oRs("amount")) 
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRs("amount"))
					dTotal = dTotal - CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal - CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					oSchema("totalamt") = oSchema("totalamt") - CDbl(oRs("amount")) 
				End If 
			Else
				oSchema.addnew 
				oSchema("accountid") = oRs("accountid")
				oSchema("accountname") = oRs("accountname")
				oSchema("accountnumber") = oRs("accountnumber")
				oSchema("ispaymentaccount") = True 
				oSchema("iscitizenaccount") = True 
				oSchema("receiptno") = oRs("paymentid")
				oSchema("paymentdate") = FormatDateTime(oRs("paymentdate"),2)
				oSchema("paymenttime") = FormatDateTime(oRs("paymentdate"),3)
				oSchema("creditamt") = 0.00
				oSchema("debitamt") = 0.00
				oSchema("totalamt") = 0.00
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					dTotal = CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					oSchema("totalamt") = oRs("amount")
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = -CDbl(oRs("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRs("amount"))
					dTotal = - CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal  - CDbl(oRs("amount"))
					oSchema("totalamt") = - CDbl(oRs("amount"))
				End If 
				iOldAccountId = CLng(oRs("accountid"))
				iOldPaymentId = CLng(oRs("paymentid"))
				oSchema("rentername") = Trim(oRs("renterfirstname") & " " & oRs("renterlastname"))
			End If 
			oSchema.Update
			oRs.MoveNext
		Loop
	End If 
	oRs.Close
	Set oRs = Nothing


	If Not bHasData Then 
		' A blank row
		oSchema.addnew 
		oSchema("accountid") = 0
		oSchema("accountname") = " "
		oSchema("accountnumber") = " "
		oSchema("receiptno") = 0
		oSchema("creditamt") = 0.00
		oSchema("debitamt") = 0.00
		oSchema("totalamt") = 0.00
		oSchema.Update
	End If 

	' Sort them 
	oSchema.Sort = "ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC, receiptno ASC"
	oSchema.MoveFirst

	' Total Row
	sTotalRow = "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td>Totals</td><td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalCredit, 2,,,0) & "</td><td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalDebit, 2,,,0) & "</td><td align=""right"" class=""moneystyle"">" & FormatNumber(dGrandTotal,2,,,0) & "</td></tr>"

	oSchema.MoveFirst

	CreateDetailExcelDownload sRptTitle, sTotalRow, sRptType

	oSchema.Close
	Set oSchema = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( ByVal iLocationid )
	Dim sSql, oRs

	sSql = "SELECT name FROM egov_class_location WHERE locationid = " & iLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetLocationName = oRs("name")
	Else
		GetLocationName = ""
	End If 

	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetAdminName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetAdminName_old( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT firstname + ' ' + lastname AS username FROM users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetAdminName = oRs("username")
	Else
		GetAdminName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' string GetJournalEntryDisplay( iJournalEntryTypeId )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryDisplay( ByVal iJournalEntryTypeId )
	Dim sSql, oRs

	sSql = "SELECt displayname FROM egov_journal_entry_types WHERE journalentrytypeid = " & iJournalEntryTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetJournalEntryDisplay = oRs("displayname") & " Only"
	Else
		GetJournalEntryDisplay = ""
	End If 

	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' void CreateDetailExcelDownload sRtpTitle, sTotalRow, sRptType
'--------------------------------------------------------------------------------------------------
Sub CreateDetailExcelDownload( ByVal sRtpTitle, ByVal sTotalRow, ByVal sRptType )
	' Pulled this in to make sub-totals

	iOldAccountId = CLng(0)
	iOldPaymentId = CLng(0)
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)
	dCreditSubTotal = CDbl(0.00)
	dDebitSubTotal = CDbl(0.00)
	dSubTotal = CDbl(0.00)

	If Not oSchema.EOF Then
		response.write "<html>"

		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"


		response.write "<body><table border=""1"">"

		' Write the title
		If sRtpTitle <> "" Then 
			response.write sRtpTitle
		End If 
		response.flush

		response.write "<tr>"
		' WRITE COLUMN HEADINGS
		response.write  "<th>Account Name</th>"
		response.write  "<th>Account Number</th>"
		response.write  "<th>Receipt No.</th>"
		response.write  "<th>Transaction Date</th>"
		response.write  "<th>Transaction Time</th>"
		response.write  "<th>Renter</th>"
		response.write  "<th>Reservation Date</th>"
		response.write  "<th>Total Amt Credited</th>"
		response.write  "<th>Total Amt Debited</th>"
		response.write  "<th>Total Amt Transfered</th>"
		response.write "</tr>"

		' WRITE DATA
		Do While Not oSchema.EOF
			If CLng(oSchema("accountid")) <> iOldAccountId Then
				If iOldAccountId <> CLng(0) And sRptType = "detail" Then 
					' Sub Total Row
					response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td>Sub-Total:</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dCreditSubTotal, 2,,,0) & "</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dDebitSubTotal, 2,,,0) & "</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal, 2,,,0) & "</td>"
					response.write "</tr>"
				End If 
				dCreditSubTotal = CDbl(0.00)
				dDebitSubTotal = CDbl(0.00)
				dSubTotal = CDbl(0.00)
				iOldAccountId = oSchema("accountid")
			End If 
			' Normal Row
			response.write "<tr>"

			' Account name
			response.write "<td align=""left"">&nbsp;" & oSchema("accountname") & "</td>"

			' Account Number
			response.write "<td align=""center"">&nbsp;" & oSchema("accountnumber") & "</td>"

			' Receipt Number
			response.write "<td align=""center"">&nbsp;" & oSchema("receiptno") & "</td>"

			' Payment Date
			response.write "<td align=""center"">" & oSchema("paymentdate") & "</td>"
			
			' Payment Date
			response.write "<td align=""center"">" & oSchema("paymenttime") & "</td>"

			' Renter
			response.write "<td align=""center"">" & oSchema("rentername") & "</td>"

			' Reservation Date
			response.write "<td align=""center"">" & oSchema("reservationdate") & "</td>"

			' Total Amt Credited
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("creditamt"),2,,,0) & "</td>"
			dCreditSubTotal = dCreditSubTotal + CDbl(oSchema("creditamt"))

			' Total Amt Debited
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("debitamt"),2,,,0) & "</td>"
			dDebitSubTotal = dDebitSubTotal + CDbl(oSchema("debitamt"))

			' Total Amt Transfered
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("totalamt"),2,,,0) & "</td>"
			dSubTotal = dSubTotal + CDbl(oSchema("totalamt"))

			response.write "</tr>"
			response.flush

			oSchema.MoveNext
		Loop
		
		If sRptType = "detail" Then 
			' Sub Total Row
			response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td></td><td></td><td>Sub-Total:</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dCreditSubTotal, 2,,,0) & "</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dDebitSubTotal, 2,,,0) & "</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal, 2,,,0) & "</td>"
			response.write "</tr>"
			response.flush
		
			' Total Row
			If sTotalRow <> "" Then 
				response.write sTotalRow
			End If 
			response.flush
		End If 

		response.write "</table></body></html>"

	End If

End Sub


%>

<!-- #include file="../export/include_excel_export.asp" -->



