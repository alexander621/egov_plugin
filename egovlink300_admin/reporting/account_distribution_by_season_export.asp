<!-- #include file="../includes/common.asp" //-->
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
	sRptTitle = "<tr>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th>Account Distribution (By Season) Summary</th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "</tr>" & vbcrlf
	sFile    = "Summary_"
	sRptType = "Summary"
Else
	sRptTitle = "<tr>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th>Account Distribution (By Season) Detail</th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "</tr>" & vbcrlf
	sFile    = "Detail_"
	sRptType = "Detail"
End If 

server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Account_Distribution_" & sFile & sDate & ".xls"

' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
fromDate = Request("fromDate")
toDate   = Request("toDate")
today    = Date()

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

If request("ClassSeasonID") <> "" Then 
	iClassSeasonID = request("ClassSeasonID")
Else 
	iClassSeasonID = ""
End If

If request("accountid") = "" Then
	iAccountNo = 0
Else
	iAccountNo = CLng(request("accountid"))
End If 


'BUILD SQL WHERE CLAUSE
sWhereClause = " AND (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
sRptTitle = sRptTitle & "<tr>" & vbcrlf
sRptTitle = sRptTitle & "    <th>Payment Date >= " & fromDate & "</th>" & vbcrlf
sRptTitle = sRptTitle & "    <th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th>" & vbcrlf
sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
sRptTitle = sRptTitle & "</tr>" & vbcrlf
sWhereClause = sWhereClause & " AND P.orgid = " & session("orgid") 

If iLocationId > 0 Then
	sWhereClause = sWhereClause & " AND adminlocationid = " & iLocationId
	sRptTitle = sRptTitle & "<tr>" & vbcrlf
	sRptTitle = sRptTitle & "    <th>Admin Location: " & GetLocationName( iLocationId )  & "</th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "</tr>" & vbcrlf
Else
	sRptTitle = sRptTitle & "<tr>" & vbcrlf
	sRptTitle = sRptTitle & "    <th>Admin Location: All Locations</th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "</tr>" & vbcrlf
End If 

If CLng(iAdminUserId) > CLng(0) Then
	sWhereClause = sWhereClause & " AND adminuserid = " & iAdminUserId
	sRptTitle = sRptTitle & "<tr>" & vbcrlf
	sRptTitle = sRptTitle & "    <th>Admin: " & GetAdminName( iAdminUserId )  & "</th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "</tr>" & vbcrlf
Else 
	sRptTitle = sRptTitle & "<tr>" & vbcrlf
    sRptTitle = sRptTitle & "    <th>Admin: All Admins</th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "</tr>" & vbcrlf
End If 

If iPaymentLocationId > 0 Then
	If iPaymentLocationId = CLng(2) Then
		sWhereClause = sWhereClause & " AND P.paymentlocationid = 3 " 
		sRptTitle = sRptTitle & "<tr>" & vbcrlf
		sRptTitle = sRptTitle & "    <th>Payment Location: Web Site</th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	   sRptTitle = sRptTitle & "</tr>" & vbcrlf
	Else
		sWhereClause = sWhereClause & " AND P.paymentlocationid < 3 " 
		sRptTitle = sRptTitle & "<tr>" & vbcrlf
		sRptTitle = sRptTitle & "    <th>Payment Location: Office</th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
		sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
       sRptTitle = sRptTitle & "</tr>" & vbcrlf
	End If 
Else
   	sRptTitle = sRptTitle & "<tr>" & vbcrlf
    sRptTitle = sRptTitle & "    <th>Payment Location: All Locations</th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
    sRptTitle = sRptTitle & "</tr>" & vbcrlf
End If 

If request("journalentrytypeid") = "" Then
	iJournalEntryTypeId = 0
Else
	iJournalEntryTypeId = CLng(request("journalentrytypeid"))
End If 

If iJournalEntryTypeId > 0 Then 
	sWhereClause = sWhereClause & " AND P.journalentrytypeid = " & iJournalEntryTypeId
	sRptTitle = sRptTitle & "<tr>" & vbcrlf
	sRptTitle = sRptTitle & "    <th>Entries: " & GetJournalEntryDisplay( iJournalEntryTypeId ) & "</th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "</tr>" & vbcrlf
Else
	sRptTitle = sRptTitle & "<tr>" & vbcrlf
	sRptTitle = sRptTitle & "    <th>Entries: Payments and Refunds</th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
'	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
'	sRptTitle = sRptTitle & "    <th></th>" & vbcrlf
	sRptTitle = sRptTitle & "</tr>" & vbcrlf
End If 

If OrgHasFeature("gl accounts") Then 
	If CLng(iAccountNo) > CLng(0) Then
		sWhereClause = sWhereClause & " AND A.accountid = " & iAccountNo & " "
		sRptTitle = sRptTitle & "<tr><th>GL Account: " & GetAccountName( iAccountNo ) & " Only</th><th></th></tr>"
	Else
		sRptTitle = sRptTitle & "<tr><th>GL Account: All GL Accounts</th><th></th><th></th></tr>"
	End If 
End If 

'Determine which season has been selected
 If iClassSeasonID <> "" Then 
    sWhereClause = sWhereClause & " AND C.ClassSeasonID = " & iClassSeasonID
 End If 

If sRptType = "Detail" Then
	DisplayDetails sWhereClause, sRptTitle
Else
	DisplaySummary sWhereClause, sRptTitle
End If 


'--------------------------------------------------------------------------------------------------
' Sub DisplaySummary( varWhereClause, sRptTitle )
'--------------------------------------------------------------------------------------------------
Sub DisplaySummary( ByVal sWhereClause, ByVal sRptTitle )
	Dim bHasData, oRequests

	iOldAccountId = CLng(0) 
	dTotal        = CDbl(0.00)
	dTotalCredit  = CDbl(0.00)
	dTotalDebit   = CDbl(0.00)
	dGrandTotal   = CDbl(0.00)
	bHasData      = False 

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	'oSchema.fields.append "accountid", adInteger, , adFldUpdatable
	oSchema.fields.append "ClassSeasonID", adInteger, , adFldUpdatable
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "creditamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "debitamt", adCurrency, , adFldUpdatable
'	oSchema.fields.append "totalamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oSchema.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount, 0 AS iscitizenaccount, "
	sSql = sSql & " sum(L.amount) as amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE A.accountid = L.accountid "
	' and (L.ispaymentaccount = 0 or (L.ispaymentaccount = 1 and L.itemid is not null and plusminus = '+')) "
	sSql = sSql & " and L.paymentid = P.paymentid "
	sSql = sSql & " AND L.amount <> 0.00 "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " GROUP BY C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount "
	sSql = sSql & " ORDER BY C.ClassSeasonID, A.accountid, L.entrytype"
	'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId Then
				If oRequests("entrytype") = "credit" Then
  					oSchema("creditamt") = oSchema("creditamt") + oRequests("amount")
  					dTotal               = dTotal + CDbl(oRequests("amount"))
  					dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
  					dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
'  					oSchema("totalamt")  = dTotal 
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRequests("amount"))
  					dTotal              = dTotal - CDbl(oRequests("amount"))
  					dGrandTotal         = dGrandTotal - CDbl(oRequests("amount"))
  					dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
'  					oSchema("totalamt") = dTotal 
				End If
			Else
				oSchema.addnew 
				'oSchema("accountid") = oRequests("accountid")
				oSchema("ClassSeasonID")    = oRequests("ClassSeasonID")
				oSchema("accountname")      = oRequests("accountname")
				oSchema("accountnumber")    = oRequests("accountnumber")
				oSchema("ispaymentaccount") = oRequests("ispaymentaccount")
				'oSchema("iscitizenaccount") = oRequests("iscitizenaccount") 
				oSchema("iscitizenaccount") = False 
				oSchema("creditamt")        = 0.00
				oSchema("debitamt")         = 0.00
'				oSchema("totalamt")         = 0.00
				If oRequests("entrytype") = "credit" Then
  					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRequests("amount"))
  					dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
  					dTotal               = CDbl(oRequests("amount"))
  					dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
'	  				oSchema("totalamt")  = oRequests("amount")
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRequests("amount"))
  					dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
  					dTotal              = -CDbl(oRequests("amount"))
  					dGrandTotal         = dGrandTotal  - CDbl(oRequests("amount"))
'  					oSchema("totalamt") = -CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
			End If 
			oSchema.Update
			oRequests.MoveNext
		Loop
	End If 
	oRequests.Close
	Set oRequests = Nothing

	'Pull the Citizen Accounts
	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount, 1 AS iscitizenaccount, "
	sSql = sSql & " sum(L.amount) as amount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, egov_organizations_to_paymenttypes OP, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE L.paymentid = P.paymentid "
	sSql = sSql & " AND L.paymenttypeid = 4 "
	sSql = sSql & " AND A.accountid = OP.accountid "
	sSql = sSql & " AND OP.paymenttypeid = L.paymenttypeid "
	sSql = sSql & " AND OP.orgid = P.orgid "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " GROUP BY C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, L.ispaymentaccount "
	sSql = sSql & " ORDER BY C.ClassSeasonID, A.accountid, L.entrytype"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId Then
				If oRequests("entrytype") = "credit" Then
 		 			oSchema("creditamt") = oRequests("amount")
 		 			dTotal               = dTotal + CDbl(oRequests("amount"))
 		 			dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
 	 				dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
' 	 				oSchema("totalamt")  = dTotal 
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = -CDbl(oRequests("amount"))
  					dTotal              = dTotal - CDbl(oRequests("amount"))
  					dGrandTotal         = dGrandTotal - CDbl(oRequests("amount"))
  					dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
'  					oSchema("totalamt") = dTotal 
				End If 
			Else
				oSchema.addnew 
				'oSchema("accountid") = oRequests("accountid")
				oSchema("ClassSeasonID")    = oRequests("ClassSeasonID")
				oSchema("accountname")      = oRequests("accountname")
				oSchema("accountnumber")    = oRequests("accountnumber")
				oSchema("ispaymentaccount") = True 
				oSchema("iscitizenaccount") = True 
				oSchema("creditamt")        = 0.00
				oSchema("debitamt")         = 0.00
'				oSchema("totalamt")         = 0.00
				If oRequests("entrytype") = "credit" Then
  					oSchema("creditamt") = CDbl(oRequests("amount"))
  					dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
  					dTotal               = CDbl(oRequests("amount"))
  					dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
'  					oSchema("totalamt")  = oRequests("amount")
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = -CDbl(oRequests("amount"))
  					dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
  					dTotal              = -CDbl(oRequests("amount"))
  					dGrandTotal         = dGrandTotal  - CDbl(oRequests("amount"))
'  					oSchema("totalamt") = -CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
			End If 
			oSchema.Update
			oRequests.MoveNext
		Loop
	End If 
	oRequests.Close
	Set oRequests = Nothing


	If Not bHasData Then 
		' A blank row
		oSchema.addnew 
		'oSchema("accountid") = 0
		oSchema("accountname")   = " "
		oSchema("accountnumber") = " "
		oSchema("creditamt")     = 0.00
		oSchema("debitamt")      = 0.00
'		oSchema("totalamt")      = 0.00
		oSchema.Update
	End If 

	' Sort them 
	oSchema.Sort = "ClassSeasonID, ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC"

	' Total Row
	sTotalRow = "<tr>" & vbcrlf
	sTotalRow = sTotalRow & "    <td></td>"      & vbcrlf
	sTotalRow = sTotalRow & "    <td></td>"      & vbcrlf
	sTotalRow = sTotalRow & "    <td>Total</td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td class=""moneystyle"">" & FormatNumber(dTotalCredit, 2) & "</td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td class=""moneystyle"">" & FormatNumber(dTotalDebit, 2)  & "</td>" & vbcrlf
	sTotalRow = sTotalRow & "</tr>" & vbcrlf

	oSchema.MoveFirst

	CreateExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub DisplayDetails( sWhereClause, sRptTitle )
'--------------------------------------------------------------------------------------------------
Sub DisplayDetails( ByVal sWhereClause, ByVal sRptTitle )
	Dim bHasData, sSql, oRequests

	iOldAccountId = CLng(0) 
	iOldPaymentId = CLng(0)
	dTotal        = CDbl(0.00)
	dTotalCredit  = CDbl(0.00)
	dTotalDebit   = CDbl(0.00)
	dGrandTotal   = CDbl(0.00)
	bHasData      = False 

	' Holding recordset
	Set oSchema = server.CreateObject("ADODB.RECORDSET")
	oSchema.fields.append "ClassSeasonID", adInteger, , adFldUpdatable
	oSchema.fields.append "accountid", adInteger, , adFldUpdatable
	oSchema.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oSchema.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oSchema.fields.append "receiptno", adInteger, , adFldUpdatable
	oSchema.fields.append "paymentdate", adDBTimeStamp, , adFldUpdatable
	oSchema.fields.append "creditamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "debitamt", adCurrency, , adFldUpdatable
'	oSchema.fields.append "totalamt", adCurrency, , adFldUpdatable
	oSchema.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oSchema.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable

	oSchema.CursorLocation = 3
	'oSchema.CursorType = 3

	oSchema.open 

	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate, "
	sSql = sSql & " L.ispaymentaccount, 0 AS iscitizenaccount "
	sSql = sSql & " FROM egov_accounts A, egov_accounts_ledger L, egov_class_payment P, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE A.accountid = L.accountid "
	' and (L.ispaymentaccount = 0 or (L.ispaymentaccount = 1 and L.itemid is not null and plusminus = '+')) "
	sSql = sSql & " AND L.paymentid = P.paymentid "
	sSql = sSql & " AND L.amount <> 0.00 "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " ORDER BY C.ClassSeasonID, A.accountid, P.paymentid, L.entrytype"
	'	response.write sSql & "<br />"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId And CLng(oRequests("paymentid")) = iOldPaymentId Then
				If oRequests("entrytype") = "credit" Then
  					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRequests("amount"))
  					dTotal               = dTotal + CDbl(oRequests("amount"))
  					dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
  					dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
'  					oSchema("totalamt")  = oSchema("totalamt") + CDbl(oRequests("amount")) 
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRequests("amount"))
  					dTotal              = dTotal - CDbl(oRequests("amount"))
  					dGrandTotal         = dGrandTotal - CDbl(oRequests("amount"))
  					dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
'  					oSchema("totalamt") = oSchema("totalamt") - CDbl(oRequests("amount")) 
				End If 
			Else
				oSchema.addnew 
				oSchema("ClassSeasonID")    = oRequests("ClassSeasonID")
				oSchema("accountid")        = oRequests("accountid")
				oSchema("accountname")      = oRequests("accountname")
				oSchema("accountnumber")    = oRequests("accountnumber")
				oSchema("ispaymentaccount") = oRequests("ispaymentaccount")
				If oRequests("accountname") = "Citizen Accounts" Then
  					oSchema("iscitizenaccount") = True 
				Else 
		  			oSchema("iscitizenaccount") = False 
				End If 
				oSchema("receiptno")   = oRequests("paymentid")
				oSchema("paymentdate") = FormatDateTime(oRequests("paymentdate"),2)
				oSchema("creditamt")   = 0.00
				oSchema("debitamt")    = 0.00
'				oSchema("totalamt")    = 0.00
				If oRequests("entrytype") = "credit" Then
  					oSchema("creditamt") = CDbl(oRequests("amount"))
  					dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
  					dTotal               = CDbl(oRequests("amount"))
  					dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
'  					oSchema("totalamt")  = oRequests("amount")
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = -CDbl(oRequests("amount"))
  					dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
  					dTotal              = - CDbl(oRequests("amount"))
  					dGrandTotal         = dGrandTotal  - CDbl(oRequests("amount"))
'  					oSchema("totalamt") = - CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
				iOldPaymentId = CLng(oRequests("paymentid"))
			End If 
			oSchema.Update
			oRequests.MoveNext
		Loop
	End If 

	oRequests.Close
	Set oRequests = Nothing

	' Citizen Accounts
	sSql = "SELECT C.ClassSeasonID, A.accountname, A.accountnumber, A.accountid, L.entrytype, P.paymentid, L.amount, P.paymentdate, "
	sSql = sSql & " ISNULL(L.paymenttypeid,0) AS paymenttypeid, P.userid, P.journalentrytypeid, L.ispaymentaccount, 1 AS iscitizenaccount "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment P, egov_accounts A, egov_organizations_to_paymenttypes OP, "
	sSql = sSql & " egov_class_list CL, egov_class_time T, egov_class C "
	sSql = sSql & " WHERE L.paymentid = P.paymentid "
	sSql = sSql & " AND L.paymenttypeid = 4 "
	sSql = sSql & " AND A.accountid = OP.accountid "
	sSql = sSql & " AND OP.paymenttypeid = L.paymenttypeid "
	sSql = sSql & " AND OP.orgid = P.orgid "
	sSql = sSql & " AND CL.classlistid = L.itemid "
	sSql = sSql & " AND CL.classtimeid = T.timeid "
	sSql = sSql & " AND C.classid = CL.classid "
	sSql = sSql & sWhereClause 
	sSql = sSql & " ORDER BY A.accountid, P.paymentid, L.entrytype"

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSql, Application("DSN"), 3, 1

	If Not oRequests.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("accountid")) = iOldAccountId And CLng(oRequests("paymentid")) = iOldPaymentId Then
				If oRequests("entrytype") = "credit" Then
  					oSchema("creditamt") = oSchema("creditamt") + CDbl(oRequests("amount"))
  					dTotal               = dTotal + CDbl(oRequests("amount"))
  					dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
  					dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
'  					oSchema("totalamt")  = oSchema("totalamt") + CDbl(oRequests("amount")) 
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = oSchema("debitamt") - CDbl(oRequests("amount"))
	  				dTotal              = dTotal - CDbl(oRequests("amount"))
		  			dGrandTotal         = dGrandTotal - CDbl(oRequests("amount"))
			  		dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
'				  	oSchema("totalamt") = oSchema("totalamt") - CDbl(oRequests("amount")) 
				End If 
			Else
				oSchema.addnew 
				oSchema("ClassSeasonID")    = oRequests("ClassSeasonID")
				oSchema("accountid")        = oRequests("accountid")
				oSchema("accountname")      = oRequests("accountname")
				oSchema("accountnumber")    = oRequests("accountnumber")
				oSchema("ispaymentaccount") = True 
				oSchema("iscitizenaccount") = True 
				oSchema("receiptno")        = oRequests("paymentid")
				oSchema("paymentdate")      = FormatDateTime(oRequests("paymentdate"),2)
				oSchema("creditamt")        = 0.00
				oSchema("debitamt")         = 0.00
'				oSchema("totalamt")         = 0.00
				If oRequests("entrytype") = "credit" Then
  					oSchema("creditamt") = CDbl(oRequests("amount"))
		  			dTotalCredit         = dTotalCredit + CDbl(oRequests("amount"))
				  	dTotal               = CDbl(oRequests("amount"))
  					dGrandTotal          = dGrandTotal + CDbl(oRequests("amount"))
'		   		oSchema("totalamt")  = oRequests("amount")
				End If 
				If oRequests("entrytype") = "debit" Then
  					oSchema("debitamt") = -CDbl(oRequests("amount"))
		  			dTotalDebit         = dTotalDebit - CDbl(oRequests("amount"))
				  	dTotal              = - CDbl(oRequests("amount"))
  					dGrandTotal         = dGrandTotal  - CDbl(oRequests("amount"))
'		  			oSchema("totalamt") = - CDbl(oRequests("amount"))
				End If 
				iOldAccountId = CLng(oRequests("accountid"))
				iOldPaymentId = CLng(oRequests("paymentid"))
			End If 
			oSchema.Update
			oRequests.MoveNext
		Loop
	End If 
	oRequests.Close
	Set oRequests = Nothing

	If Not bHasData Then 
		' A blank row
		oSchema.addnew 
		oSchema("accountid")     = 0
		oSchema("accountname")   = " "
		oSchema("accountnumber") = " "
		oSchema("receiptno")     = 0
		oSchema("creditamt")     = 0.00
		oSchema("debitamt")      = 0.00
'		oSchema("totalamt")      = 0.00
		oSchema.Update
	End If 

	' Sort them 
	oSchema.Sort = "ClassSeasonID, ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC, receiptno ASC"
	oSchema.MoveFirst

	' Total Row
	sTotalRow = "<tr>" & vbcrlf
	sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td></td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td>Totals</td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td class=""moneystyle"">" & FormatNumber(dTotalCredit, 2) & "</td>" & vbcrlf
	sTotalRow = sTotalRow & "    <td class=""moneystyle"">" & FormatNumber(dTotalDebit, 2)  & "</td>" & vbcrlf
	' sTotalRow = sTotalRow & "    <td>" & FormatNumber(dGrandTotal,2)   & "</td>" & vbcrlf

	sTotalRow = sTotalRow & "</tr>" & vbcrlf

	oSchema.MoveFirst

	CreateDetailExcelDownload sRptTitle, sTotalRow

	oSchema.Close
	Set oSchema = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Function GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName( ByVal iLocationid )
	Dim sSql, oLocation

	sSql = "SELECT name FROM egov_class_location WHERE locationid = " & iLocationId

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSql, Application("DSN"), 3, 1
	
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
Function GetAdminName( ByVal iUserId )
	Dim sSql, oName

	sSql = "SELECT firstname + ' ' + lastname AS username FROM users WHERE userid = " & iUserId 

	Set oName = Server.CreateObject("ADODB.Recordset")
	oName.Open sSql, Application("DSN"), 3, 1

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
Function GetJournalEntryDisplay( ByVal iJournalEntryTypeId )
	Dim sSql, oType

	sSql = "SELECT displayname FROM egov_journal_entry_types WHERE journalentrytypeid = " & iJournalEntryTypeId

	Set oType = Server.CreateObject("ADODB.Recordset")
	oType.Open sSql, Application("DSN"), 3, 1
	
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
Sub CreateDetailExcelDownload( ByVal sRtpTitle, ByVal sTotalRow )
	' Pulled this in to make sub-totals

	iOldAccountId   = CLng(0)
	iOldPaymentId   = CLng(0)
	dTotal          = CDbl(0.00)
	dTotalCredit    = CDbl(0.00)
	dTotalDebit     = CDbl(0.00)
	dGrandTotal     = CDbl(0.00)
	dCreditSubTotal = CDbl(0.00)
	dDebitSubTotal  = CDbl(0.00)
	dSubTotal       = CDbl(0.00)

	If Not oSchema.EOF Then
		response.write "<html><body>"
		
		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"
		
		response.write vbcrlf & "<table border=""0"">"

		' Write the title
		If sRtpTitle <> "" Then 
			response.write sRtpTitle
		End If 
		response.flush

		response.write "<tr>" & vbcrlf
		' WRITE COLUMN HEADINGS
		For Each fldLoop in oSchema.Fields
   			If fldLoop.Name <> "accountid" And fldLoop.Name <> "ispaymentaccount" And fldLoop.Name <> "iscitizenaccount" Then 
  	   			response.write  "<th>" & fldLoop.Name & "</th>" & vbcrlf
   			End If 
		Next

		response.write "</tr>" & vbcrlf
		response.flush

		' WRITE DATA
		Do While Not oSchema.EOF
			If CLng(oSchema("accountid")) <> iOldAccountId Then
				If iOldAccountId <> CLng(0) Then 
					' Sub Total Row
					response.write "<tr>" & vbcrlf
					response.write "    <td></td>" & vbcrlf
					response.write "    <td></td>" & vbcrlf
					response.write "    <td></td>" & vbcrlf
					response.write "    <td></td>" & vbcrlf
					response.write "    <td>Sub-Total:</td>" & vbcrlf
					response.write "    <td class=""moneystyle"">" & FormatNumber(dCreditSubTotal, 2) & "</td>" & vbcrlf
					response.write "    <td class=""moneystyle"">" & FormatNumber(-dDebitSubTotal, 2) & "</td>" & vbcrlf
					response.write "</tr>" & vbcrlf
					response.flush
				End If 
				dCreditSubTotal = CDbl(0.00)
				dDebitSubTotal  = CDbl(0.00)
				dSubTotal       = CDbl(0.00)
				iOldAccountId   = oSchema("accountid")
			End If 
			' Normal Row
			response.write "<tr>" & vbcrlf
			For Each fldLoop in oSchema.Fields
				sFieldValue = trim(fldLoop.Value)
				
				' REMOVE LINE BREAKS
				If Not IsNull(sFieldValue) Then
					sFieldValue = Replace(sFieldValue,Chr(10),"")
					sFieldValue = Replace(sFieldValue,Chr(13),"")
				End If
				
				If fldLoop.Name = "creditamt" Then
					dCreditSubTotal = dCreditSubTotal + CDbl(sFieldValue)
					dSubTotal = dSubTotal + CDbl(sFieldValue)
				End If 
				If fldLoop.Name = "debitamt" Then
					dDebitSubTotal = dDebitSubTotal - CDbl(sFieldValue)
					dSubTotal = dSubTotal + CDbl(sFieldValue)
				End If 

				If fldLoop.Name <> "accountid" And fldLoop.Name <> "ispaymentaccount" And fldLoop.Name <> "iscitizenaccount" Then
					If UCase(fldLoop.Name) = "CLASSSEASONID" Then 
						response.write "<td>" & getSeasonName(sFieldValue) & "</td>" & vbcrlf
					Else 
       					response.write "<td"
						If fldLoop.Type = 6 Then 
							' This type is currency
							response.write " class=""moneystyle"""
						End If 
						response.write ">" & sFieldValue & "</td>" & vbcrlf
					End If 
				End If 
			Next
			response.write "</tr>" & vbcrlf
			response.flush
			 

			oSchema.MoveNext
		Loop
		
		' Sub Total Row
		response.write "<tr>" & vbcrlf
		response.write "    <td></td>" & vbcrlf
		response.write "    <td></td>" & vbcrlf
		response.write "    <td></td>" & vbcrlf
		response.write "    <td></td>" & vbcrlf
		response.write "    <td>Sub-Total:</td>" & vbcrlf
		response.write "    <td class=""moneystyle"">" & FormatNumber(dCreditSubTotal, 2) & "</td>" & vbcrlf
		response.write "    <td class=""moneystyle"">" & FormatNumber(-dDebitSubTotal, 2) & "</td>" & vbcrlf
		response.write "</tr>" & vbcrlf
		response.flush

		' Total Row
		If sTotalRow <> "" Then 
			response.write sTotalRow
		End If 
		response.flush

		response.write "</table></body></html>"
	Else

		' NO DATA

	End If

End Sub


'---------------------------------------------------------
function getSeasonName( ByVal p_classseasonid )
	Dim lcl_return, sSql, oSeasonName

	lcl_return = ""

	If p_classseasonid <> "" Then 
		sSql = "SELECT seasonname FROM egov_class_seasons WHERE classseasonid = " & p_classseasonid

		Set oSeasonName = Server.CreateObject("ADODB.Recordset")
		oSeasonName.Open sSql, Application("DSN"), 3, 1

		If Not oSeasonName.eof Then 
		lcl_return = oSeasonName("seasonname")
		End If 

		oSeasonName.close
		Set oSeasonName = Nothing 

	End If 

	getSeasonName = lcl_return

End Function 


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
<!-- #include file="../export/include_excel_export.asp" -->

