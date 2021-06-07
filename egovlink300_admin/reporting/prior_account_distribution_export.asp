<%
	Dim sSql, oRequests, oSchema, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal
	Dim iLocationId, toDate, fromDate, sDateRange, iPaymentLocationId, iReportType, sAdminlocation
	Dim sFile, sRptTitle, sWhereClause, iJournalEntryTypeId

	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	sWhereClause = ""

	If request("reporttype") = "" Then 
		iReportType = CLng(1)
	Else
		iReportType = CLng(request("reporttype"))
	End If 

	If iReportType = CLng(1) Then
		sRptTitle = "<tr><th></th><th>Account Distribution Summary</th><th></th><th></th><th></th></tr>"
		sFile = "Summary_"
		sRptType = "Summary"
	Else
		sRptTitle = "<tr><th></th><th>Account Distribution Detail</th><th></th><th></th><th></th></tr>"
		sFile = "Detail_"
		sRptType = "Detail"
	End If 

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

	' BUILD SQL WHERE CLAUSE
	sWhereClause = " AND (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
	sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & fromDate & "</th><th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th></tr>"
	sWhereClause = sWhereClause & " AND P.orgid = " & session("orgid") 

	If iLocationId > 0 Then
		sWhereClause = sWhereClause & " AND adminlocationid = " & iLocationId
		sRptTitle = sRptTitle & "<tr><th>Admin Location: " & GetLocationName( iLocationId )  & "</th><th></th><th></th><th></th><th></th></tr>"
	Else
		sRptTitle = sRptTitle & "<tr><th>Admin Location: All Locations</th><th></th><th></th><th></th><th></th></tr>"
	End If 

	If CLng(iAdminUserId) > CLng(0) Then
		sWhereClause = sWhereClause & " AND adminuserid = " & iAdminUserId
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

	If iJournalEntryTypeId > 0 Then 
		sWhereClause = sWhereClause & " AND P.journalentrytypeid = " & iJournalEntryTypeId
		sRptTitle = sRptTitle & "<tr><th>Entries: " & GetJournalEntryDisplay( iJournalEntryTypeId ) & "</th><th></th><th></th><th></th><th></th></tr>"
	Else
		sRptTitle = sRptTitle & "<tr><th>Entries: Payments and Refunds</th><th></th><th></th><th></th><th></th></tr>"
	End If 

	

	If sRptType = "Detail" Then
		DisplayDetails sWhereClause, sRptTitle
	Else
		DisplaySummary sWhereClause, sRptTitle
	End If 




'--------------------------------------------------------------------------------------------------
' Sub DisplaySummary( varWhereClause, sRptTitle )
'--------------------------------------------------------------------------------------------------
Sub DisplaySummary( sWhereClause, sRptTitle )
	iOldAccountId = CLng(0) 
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	'oSchema.fields.append "accountid", adInteger, , adFldUpdatable
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "creditamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "debitamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "totalamt", adCurrency, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, sum(L.amount) as amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P "
	sSql = sSql & " WHERE A.accountid = L.accountid and (L.ispaymentaccount = 0 or (L.ispaymentaccount = 1 and L.itemid is not null and plusminus = '+')) "
	sSql = sSql & " and L.paymentid = P.paymentid and L.amount <> 0.00 " & sWhereClause 
	sSql = sSql & " GROUP BY A.accountname, A.accountnumber, A.accountid, L.entrytype ORDER BY A.accountid, L.entrytype"
'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	If Not oRequests.EOF Then

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId Then
				If oRequests("entrytype") = "credit" Then
					oSchema("creditamt") = oRequests("amount")
					dTotal = dTotal + CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					oSchema("totalamt") = dTotal 
				End If 
				If oRequests("entrytype") = "debit" Then
					oSchema("debitamt") = -CDbl(oRequests("amount"))
					dTotal = dTotal - CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal - CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					oSchema("totalamt") = dTotal 
				End If 
			Else
				oSchema.addnew 
				'oSchema("accountid") = oRequests("accountid")
				oSchema("accountname") = oRequests("accountname")
				oSchema("accountnumber") = oRequests("accountnumber")
				oSchema("creditamt") = 0.00
				oSchema("debitamt") = 0.00
				oSchema("totalamt") = 0.00
				If oRequests("entrytype") = "credit" Then
					oSchema("creditamt") = CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					dTotal = CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
					oSchema("totalamt") = oRequests("amount")
				End If 
				If oRequests("entrytype") = "debit" Then
					oSchema("debitamt") = -CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					dTotal = -CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal  - CDbl(oRequests("amount"))
					oSchema("totalamt") = -CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
			End If 
			oSchema.Update
			oRequests.MoveNext
		Loop
	Else
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
	oSchema.Sort = "accountname ASC, accountnumber ASC "

	' Total Row
	sTotalRow = "<tr><td></td><td>Total</td><td>" & FormatNumber(dTotalCredit, 2) & "</td><td>" & FormatNumber(dTotalDebit, 2) & "</td><td>" & FormatNumber(dGrandTotal,2) & "</td></tr>"

	oSchema.MoveFirst

	CreateExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

	oRequests.Close
	Set oRequests = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub DisplayDetails( sWhereClause, sRptTitle )
'--------------------------------------------------------------------------------------------------
Sub DisplayDetails( sWhereClause, sRptTitle )
	iOldAccountId = CLng(0) 
	iOldPaymentId = CLng(0)
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	oSchema.fields.append "accountid", adInteger, , adFldUpdatable
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "receiptno", adInteger, , adFldUpdatable
	oSchema.fields.append "paymentdate", adDBTimeStamp, , adFldUpdatable
	oSchema.fields.append "creditamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "debitamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "totalamt", adCurrency, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

	sSql = "SELECT A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P "
	sSql = sSql & " WHERE A.accountid = L.accountid and (L.ispaymentaccount = 0 or (L.ispaymentaccount = 1 and L.itemid is not null and plusminus = '+')) "
	sSql = sSql & " and L.paymentid = P.paymentid and L.amount <> 0.00 " & sWhereClause 
	sSql = sSql & " ORDER BY A.accountid, P.paymentid, L.entrytype"
'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	If Not oRequests.EOF Then

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId And CLng(oRequests("paymentid")) = iOldPaymentId Then
				If oRequests("entrytype") = "credit" Then
					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRequests("amount"))
					dTotal = dTotal + CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					oSchema("totalamt") = oSchema("totalamt") + CDbl(oRequests("amount")) 
				End If 
				If oRequests("entrytype") = "debit" Then
					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRequests("amount"))
					dTotal = dTotal - CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal - CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					oSchema("totalamt") = oSchema("totalamt") - CDbl(oRequests("amount")) 
				End If 
			Else
				oSchema.addnew 
				oSchema("accountid") = oRequests("accountid")
				oSchema("accountname") = oRequests("accountname")
				oSchema("accountnumber") = oRequests("accountnumber")
				oSchema("receiptno") = oRequests("paymentid")
				oSchema("paymentdate") = FormatDateTime(oRequests("paymentdate"),2)
				oSchema("creditamt") = 0.00
				oSchema("debitamt") = 0.00
				oSchema("totalamt") = 0.00
				If oRequests("entrytype") = "credit" Then
					oSchema("creditamt") = CDbl(oRequests("amount"))
					dTotalCredit = dTotalCredit + CDbl(oRequests("amount"))
					dTotal = CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
					oSchema("totalamt") = oRequests("amount")
				End If 
				If oRequests("entrytype") = "debit" Then
					oSchema("debitamt") = -CDbl(oRequests("amount"))
					dTotalDebit = dTotalDebit - CDbl(oRequests("amount"))
					dTotal = - CDbl(oRequests("amount"))
					dGrandTotal = dGrandTotal  - CDbl(oRequests("amount"))
					oSchema("totalamt") = - CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
				iOldPaymentId = CLng(oRequests("paymentid"))
			End If 
			oSchema.Update
			oRequests.MoveNext
		Loop
	Else
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
	oSchema.Sort = "accountname ASC, accountnumber ASC, receiptno ASC"

	' Total Row
	sTotalRow = "<tr><td></td><td></td><td></td><td>Totals</td><td>" & FormatNumber(dTotalCredit, 2) & "</td><td>" & FormatNumber(dTotalDebit, 2) & "</td><td>" & FormatNumber(dGrandTotal,2) & "</td></tr>"

	oSchema.MoveFirst

	CreateDetailExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

	oRequests.Close
	Set oRequests = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( iLocationid )
	Dim sSql, oLocation

	sSql = "Select name from egov_class_location where locationid = " & iLocationId

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 3, 1
	
	If Not oLocation.EOF Then 
		GetLocationName = oLocation("name")
	Else
		GetLocationName = ""
	End If 

	oLocation.Close 
	Set oLocation = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetAdminName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetAdminName( iUserId )
	Dim sSql, oName

	sSql = "SELECT firstname + ' ' + lastname as username FROM users Where userid = " & iUserId 

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSQL, Application("DSN"), 3, 1

	If Not oName.EOF Then
		GetAdminName = oName("username")
	Else
		GetAdminName = ""
	End If 

	oName.close
	Set oName = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetJournalEntryDisplay( iJournalEntryTypeId )
'--------------------------------------------------------------------------------------------------
Function GetJournalEntryDisplay( iJournalEntryTypeId )
	Dim sSql, oType

	sSql = "Select displayname from egov_journal_entry_types where journalentrytypeid = " & iJournalEntryTypeId

	Set oType = Server.CreateObject("ADODB.Recordset")
	oType.Open sSQL, Application("DSN"), 3, 1
	
	If Not oType.EOF Then 
		GetJournalEntryDisplay = oType("displayname") & " Only"
	Else
		GetJournalEntryDisplay = ""
	End If 

	oType.Close 
	Set oType = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub CreateDetailExcelDownload( sRtpTitle, sTotalRow )
'--------------------------------------------------------------------------------------------------
Sub CreateDetailExcelDownload( sRtpTitle, sTotalRow )
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

	If NOT oSchema.EOF Then
		response.write "<html><body><table border=""1"">"

		' Write the title
		If sRtpTitle <> "" Then 
			response.write sRtpTitle
		End If 

		response.write "<tr>"
		' WRITE COLUMN HEADINGS
		For Each fldLoop in oSchema.Fields
			If fldLoop.Name <> "accountid" Then 
				response.write  "<th>" & fldLoop.Name & "</th>"
			End If 
		Next
		response.write "</tr>"
		response.flush

		' WRITE DATA
		Do While NOT oSchema.EOF
			If CLng(oSchema("accountid")) <> iOldAccountId Then
				If iOldAccountId <> CLng(0) Then 
					' Sub Total Row
					response.write vbcrlf & "<tr><td></td><td></td><td></td><td>Sub-Total:</td>"
					response.write "<td>" & FormatNumber(dCreditSubTotal, 2) & "</td>"
					response.write "<td>" & FormatNumber(-dDebitSubTotal, 2) & "</td>"
					response.write "<td>" & FormatNumber(dSubTotal, 2) & "</td>"
					response.write "</tr>"
					response.flush
				End If 
				dCreditSubTotal = CDbl(0.00)
				dDebitSubTotal = CDbl(0.00)
				dSubTotal = CDbl(0.00)
				iOldAccountId = oSchema("accountid")
			End If 
			' Normal Row
			response.write "<tr>"
			For Each fldLoop in oSchema.Fields
				sFieldValue = trim(fldLoop.Value)
				
				' REMOVE LINE BREAKS
				If NOT ISNULL(sFieldValue) Then
					sFieldValue = replace(sFieldValue,chr(10),"")
					sFieldValue = replace(sFieldValue,chr(13),"")
				End If
				
				If fldLoop.Name = "creditamt" Then
					dCreditSubTotal = dCreditSubTotal + CDbl(sFieldValue)
					dSubTotal = dSubTotal + CDbl(sFieldValue)
				End If 
				If fldLoop.Name = "debitamt" Then
					dDebitSubTotal = dDebitSubTotal - CDbl(sFieldValue)
					dSubTotal = dSubTotal + CDbl(sFieldValue)
				End If 

				If fldLoop.Name <> "accountid" Then
					response.write "<td>" & sFieldValue & "</td>"
				End If 
			Next
			response.write "</tr>"
			response.flush
			 

			oSchema.MoveNext
		Loop
		
		' Sub Total Row
		response.write vbcrlf & "<tr><td></td><td></td><td></td><td>Sub-Total:</td>"
		response.write "<td>" & FormatNumber(dCreditSubTotal, 2) & "</td>"
		response.write "<td>" & FormatNumber(-dDebitSubTotal, 2) & "</td>"
		response.write "<td>" & FormatNumber(dSubTotal, 2) & "</td>"
		response.write "</tr>"
		response.flush

		' Total Row
		If sTotalRow <> "" Then 
			response.write sTotalRow
		End If 

		response.write "</table></body></html>"
		response.flush
	Else

		' NO DATA

	End If

End Sub


%>

<!-- #include file="../export/include_excel_export.asp" -->

<!-- #include file="../includes/adovbs.inc" -->

