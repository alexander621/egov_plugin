<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: citizen_account_distribution_export.asp
' AUTHOR: Steve Loar
' CREATED: 01/30/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/30/2013	Steve Loar - INITIAL VERSION
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim sSql, oRs, oSchema, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal
Dim iLocationId, toDate, fromDate, sDateRange, iPaymentLocationId, iReportType, sAdminlocation
Dim sFile, sRptTitle, sWhereClause, iJournalEntryTypeId
Dim from_time, to_time, where_time

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
sWhereClause = ""

If request("reporttype") = "" Then 
	iReportType = CLng(1)
Else
	iReportType = CLng(request("reporttype"))
End If 

If iReportType = CLng(1) Then
	sRptTitle = "<tr><th></th><th>Citizen Account Distribution Summary</th><th></th><th></th><th></th><th></th></tr>"
	sFile = "Summary_"
	sRptType = "Summary"
Else
	sRptTitle = "<tr><th></th><th>Citizen Account Distribution Detail</th><th></th><th></th><th></th><th></th></tr>"
	sFile = "Detail_"
	sRptType = "Detail"
End If 

server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=CitizenAccountDistribution_" & sFile & ".xls"

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

If request("adminuserid") = "" Then
	iAdminUserId = 0
Else
	iAdminUserId = CLng(request("adminuserid"))
End If 

If request("accountid") = "" Then
	iAccountNo = 0
Else
	iAccountNo = CLng(request("accountid"))
End If 


' BUILD SQL WHERE CLAUSE
sWhereClause = " AND P.orgid = " & session("orgid") 

'sWhereClause = sWhereClause & " AND (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
'sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & fromDate & "</th><th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th></tr>"
If from_time = "none" Then 
	sWhereClause = sWhereClause & " AND paymentDate >= '" & fromDate & "' "
	sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & fromDate & "</th>"
Else
	where_time = CDate( fromdate & " " & from_time )
	sWhereClause = sWhereClause & " AND paymentDate >= '" & where_time & "' "
	sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & where_time & "</th>"
End If 

If to_time = "none" Then 
	sWhereClause = sWhereClause & " AND paymentDate <= '" & DateAdd("d",1,toDate) & "' "
	sRptTitle = sRptTitle & "<th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th><th></th></tr>"
Else 
	where_time = CDate( todate & " " & to_time )
	sWhereClause = sWhereClause & " AND paymentDate <= '" & where_time & "' "
	sRptTitle = sRptTitle & "<th>AND Payment Date <= " & where_time & "</th><th></th><th></th><th></th><th></th></tr>"
End If 

If iLocationId > 0 Then
	sWhereClause = sWhereClause & " AND adminlocationid = " & iLocationId
	sRptTitle = sRptTitle & "<tr><th>Admin Location: " & GetLocationName( iLocationId )  & "</th><th></th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Admin Location: All Locations</th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If CLng(iAdminUserId) > CLng(0) Then
	sWhereClause = sWhereClause & " AND adminuserid = " & iAdminUserId
	sRptTitle = sRptTitle & "<tr><th>Admin: " & GetAdminName( iAdminUserId )  & "</th><th></th><th></th><th></th><th></th><th></th></tr>"
Else 
	sRptTitle = sRptTitle & "<tr><th>Admin: All Admins</th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If request("journalentrytypeid") = "" Then
	iJournalEntryTypeId = 0
Else
	iJournalEntryTypeId = CLng(request("journalentrytypeid"))
End If 

If iJournalEntryTypeId > 0 Then 
	sWhereClause = sWhereClause & " AND P.journalentrytypeid = " & iJournalEntryTypeId
	sRptTitle = sRptTitle & "<tr><th>Entries: " & GetJournalEntryDisplay( iJournalEntryTypeId ) & "</th><th></th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Entries: Payments and Refunds</th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If OrgHasFeature("gl accounts") Then 
	If CLng(iAccountNo) > CLng(0) Then
		sWhereClause = sWhereClause & " AND A.accountid = " & iAccountNo & " "
		sRptTitle = sRptTitle & "<tr><th>GL Account: " & GetAccountName( iAccountNo ) & " Only</th><th></th><th></th><th></th><th></th></tr>"
	Else
		sRptTitle = sRptTitle & "<tr><th>GL Account: All GL Accounts</th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 
End If 

If sRptType = "Detail" Then
	DisplayDetails sWhereClause, sRptTitle
Else
	DisplaySummary sWhereClause, sRptTitle
End If 


'--------------------------------------------------------------------------------------------------
' DisplaySummary varWhereClause, sRptTitle 
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

	' Pull the Citizen Accounts
	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount, 1 AS iscitizenaccount, sum(L.amount) AS amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_organizations_to_paymenttypes OP "
	sSql = sSql & " WHERE L.paymentid = P.paymentid AND P.isforrentals = 0 AND P.journalentrytypeid > 2 AND A.accountid = OP.accountid AND "
	sSql = sSql & " OP.paymenttypeid = L.paymenttypeid AND OP.orgid = P.orgid " & sWhereClause 
	sSql = sSql & " GROUP BY A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount ORDER BY A.accountid, L.entrytype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) = iOldAccountId Then
				If oRs("entrytype") = "credit" Then
					oSchema("creditamt") = oRs("amount")
					dTotal = dTotal + CDbl(oRs("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRs("amount"))
					oSchema("totalamt") = dTotal 
				End If 
				If oRs("entrytype") = "debit" Then
					oSchema("debitamt") = -CDbl(oRs("amount"))
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


	If Not bHasData Then 
		' A blank row
		oSchema.addnew 
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
	sTotalRow = "<tr><td></td><td>Total:</td>"
	sTotalRow = sTotalRow & "<td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalCredit, 2) & "</td>"
	sTotalRow = sTotalRow & "<td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalDebit, 2) & "</td>"
	sTotalRow = sTotalRow & "<td align=""right"" class=""moneystyle"">" & FormatNumber(dGrandTotal,2) & "</td></tr>"

	oSchema.MoveFirst

	CreateSummaryExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' DisplayDetails sWhereClause, sRptTitle 
'--------------------------------------------------------------------------------------------------
Sub DisplayDetails( ByVal sWhereClause, ByVal sRptTitle )
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

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

	' Citizen Accounts
	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate, "
	sSql = sSql & " ISNULL(L.paymenttypeid,0) AS paymenttypeid, ISNULL(P.userid,0) AS userid, P.journalentrytypeid, L.ispaymentaccount, 1 AS iscitizenaccount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment P, egov_accounts A, egov_organizations_to_paymenttypes OP "
	sSql = sSql & " WHERE L.paymentid = P.paymentid AND P.isforrentals = 0 AND P.journalentrytypeid > 2 AND "
	sSql = sSql & " A.accountid = OP.accountid AND OP.paymenttypeid = L.paymenttypeid AND OP.orgid = P.orgid " & sWhereClause 
	sSql = sSql & " ORDER BY A.accountid, P.paymentid, L.entrytype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

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
	sTotalRow = "<tr><td></td><td></td><td></td><td></td><td>Total:</td>"
	sTotalRow = sTotalRow & "<td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalCredit, 2) & "</td>"
	sTotalRow = sTotalRow & "<td align=""right"" class=""moneystyle"">" & FormatNumber(dTotalDebit, 2) & "</td>"
	sTotalRow = sTotalRow & "<td align=""right"" class=""moneystyle"">" & FormatNumber(dGrandTotal,2) & "</td></tr>"

	oSchema.MoveFirst

	CreateDetailExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string sLocation = GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( ByVal iLocationid )
	Dim sSql, oRs

	sSql = "SELECT name FROM egov_class_location WHERE locationid = " & iLocationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetLocationName = oRs("name")
	Else
		GetLocationName = ""
	End If 

	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string sName = GetAdminName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetAdminName( ByVal iUserId )
	Dim sSql, oRs

	sSql = "SELECT firstname + ' ' + lastname AS username FROM users WHERE userid = " & iUserId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetAdminName = oRs("username")
	Else
		GetAdminName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' string sDisplay = GetJournalEntryDisplay( iJournalEntryTypeId )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryDisplay( ByVal iJournalEntryTypeId )
	Dim sSql, oRs

	sSql = "SELECt displayname FROM egov_journal_entry_types WHERE journalentrytypeid = " & iJournalEntryTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then 
		GetJournalEntryDisplay = oRs("displayname") & " Only"
	Else
		GetJournalEntryDisplay = ""
	End If 

	oRs.Close 
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' CreateDetailExcelDownload sRtpTitle, sTotalRow 
'--------------------------------------------------------------------------------------------------
Sub CreateDetailExcelDownload( ByVal sRtpTitle, ByVal sTotalRow )
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
		response.write vbcrlf & " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"
		
		response.write vbcrlf & "<body><table border=""1"">"

		' Write the title
		If sRtpTitle <> "" Then 
			response.write sRtpTitle
		End If 
		response.flush

		response.write "<tr>"
		' WRITE COLUMN HEADINGS
		response.write "<th>Account Name</th>"
		response.write "<th>Account Number</th>"
		response.write "<th>Receipt No.</th>"
		response.write "<th>Date</th>"
		response.write "<th>Time</th>"
		response.write "<th>Total Amount Credited</th>"
		response.write "<th>Total Amount Debited</th>"
		response.write "<th>Total Amount Transferred</th>"
		response.write "</tr>"
		response.flush

		' WRITE DATA
		Do While Not oSchema.EOF
			If CLng(oSchema("accountid")) <> iOldAccountId Then
				If iOldAccountId <> CLng(0) Then 
					' Sub Total Row
					response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td>Sub-Total:</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dCreditSubTotal, 2) & "</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(-dDebitSubTotal, 2) & "</td>"
					response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal, 2) & "</td>"
					response.write "</tr>"
				End If 
				dCreditSubTotal = CDbl(0.00)
				dDebitSubTotal = CDbl(0.00)
				dSubTotal = CDbl(0.00)
				iOldAccountId = oSchema("accountid")
			End If 

			' Normal Row
			response.write "<tr>"

			' Account Name
			response.write "<td align=""left"">&nbsp;" & oSchema("accountname") & "</td>"

			' Account Number
			response.write "<td align=""center"">&nbsp;" & oSchema("accountnumber") & "</td>"

			' Receipt Number
			response.write "<td align=""center"">&nbsp;" & oSchema("receiptno") & "</td>"

			' Transaction Date
			response.write "<td align=""center"">&nbsp;" & oSchema("paymentdate") & "</td>"
			
			' Transaction Time
			response.write "<td align=""center"">&nbsp;" & oSchema("paymenttime") & "</td>"

			' Credit Amount
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("creditamt"),2,,,0) & "</td>"
			dCreditSubTotal = dCreditSubTotal + CDbl(oSchema("creditamt"))
			dSubTotal = dSubTotal + CDbl(oSchema("creditamt"))

			' Debit Amount
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("debitamt"),2,,,0) & "</td>"
			dDebitSubTotal = dDebitSubTotal - CDbl(oSchema("debitamt"))
			dSubTotal = dSubTotal + CDbl(oSchema("debitamt"))

			' Total Amount
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("totalamt"),2,,,0) & "</td>"

			response.write "</tr>"
			response.flush

			oSchema.MoveNext
		Loop
		
		' Final Sub Total Row
		response.write vbcrlf & "<tr><td></td><td></td><td></td><td></td><td>Sub-Total:</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dCreditSubTotal, 2) & "</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(-dDebitSubTotal, 2) & "</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(dSubTotal, 2) & "</td>"
		response.write "</tr>"
		response.flush

		' Total Row
		If sTotalRow <> "" Then 
			response.write sTotalRow
		End If 
		response.flush

		response.write "</table></body></html>"

	End If

End Sub


'--------------------------------------------------------------------------------------------------
' CreateSummaryExcelDownload sRtpTitle, sTotalRow 
'--------------------------------------------------------------------------------------------------
Sub CreateSummaryExcelDownload( ByVal sRtpTitle, ByVal sTotalRow )

	If Not oSchema.EOF Then
		response.write "<html>"
		
		response.write vbcrlf & "<style>  "
		response.write vbcrlf & " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"
		
		response.write vbcrlf & "<body><table border=""1"">"

		' Write the title
		If sRtpTitle <> "" Then 
			response.write sRtpTitle
		End If 

		' WRITE COLUMN HEADINGS
		response.write vbcrlf & "<tr>"
		response.write "<th>Account Name</th>"
		response.write "<th>Account Number</th>"
		response.write "<th>Total Amount Credited</th>"
		response.write "<th>Total Amount Debited</th>"
		response.write "<th>Total Amount Transferred</th>"
		response.write "</tr>"
		response.flush

		Do While Not oSchema.EOF
			' Normal Row
			response.write "<tr>"

			' Account Name
			response.write "<td align=""left"">&nbsp;" & oSchema("accountname") & "</td>"

			' Account Number
			response.write "<td align=""center"">&nbsp;" & oSchema("accountnumber") & "</td>"

			' Credit Amount
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("creditamt"),2,,,0) & "</td>"

			' Debit Amount
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("debitamt"),2,,,0) & "</td>"

			' Total Amount
			response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(oSchema("totalamt"),2,,,0) & "</td>"

			response.write "</tr>"
			response.flush

			oSchema.MoveNext
		Loop

		' Total Row
		If sTotalRow <> "" Then 
			response.write sTotalRow
		End If 

		response.write "</table></body></html>"

	End If

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetAccountName( iAccountId )
'--------------------------------------------------------------------------------------------------
Function GetAccountName( ByVal iAccountId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountname,'') AS accountname FROM egov_accounts "
	sSql = sSql & "WHERE accountid = " & iAccountId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetAccountName = oRs("accountname")
	Else
		GetAccountName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


%>




