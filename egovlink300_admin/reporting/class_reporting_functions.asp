<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: class_reporting_functions.asp
' AUTHOR: SteveLoar
' CREATED: 03/24/2014
' COPYRIGHT: Copyright 2014 eclink, inc.
'			 All Rights Reserved.
'
' Description:  The data pulls for class financial reports
'
' MODIFICATION HISTORY
' 1.0   03/24/20014	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
 
'------------------------------------------------------------------------------------------------------------
' void Display_Class_Payment_Report sWhereClause 
'------------------------------------------------------------------------------------------------------------
Sub Display_Class_Payment_Report( ByVal sWhereClause )
	Dim sSql, oRs, oDisplay, iOldPaymentId, dCashTotal, dCheckTotal, dCardtotal, dOtherTotal, dMemoTotal
	Dim dGrandTotal, dCCCTotal, dCCCSubTotal, bHasData

	iOldPaymentId = CLng(0) 
	dCCCTotal = CDbl(0.0)
	bHasData = False 

	' make a holding recordset
	Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
	oDisplay.fields.append "paymentid", adInteger, , adFldUpdatable
	oDisplay.fields.append "paymentdate", adVariant, 10, adFldUpdatable
	oDisplay.fields.append "item", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userid", adInteger, , adFldUpdatable
	oDisplay.fields.append "userfname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userlname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "userhomephone", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "checkamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "checkno", adVarChar, 20, adFldUpdatable
	oDisplay.fields.append "cashamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "cardamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "cccsubtotal", adCurrency, , adFldUpdatable
	oDisplay.fields.append "otheramt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "memoamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "paymenttotal", adCurrency, , adFldUpdatable

	oDisplay.CursorLocation = 3
	'oDisplay.CursorType = 3
	oDisplay.open 

	' Pull Class Purchases
	bHasData = PullClassPaymentData( "egov_class_to_payment_method", sWhereClause, oDisplay )
	'If PullClassPaymentData( "egov_class_to_payment_method", sWhereClause, oDisplay ) Then 
	'	bHasData = True
	'End If 


	If bHasData Then 
		' Sort the data by paymentid
		oDisplay.sort = "paymentid"
		' Show results
		oDisplay.MoveFirst

		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Receipt</th><th>Date</th><th>Payee</th>"
		response.write "<th>Check Amt<br />Check #</th><th>Cash Amt</th><th>Card Amt</th><th>Total Chck<br />Cash, Card</th><th>Other Amt</th>"
		response.write "<th>Memo Amt</th><th>Total<br />Paid</th></tr>"
		bgcolor = "#eeeeee"
		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			If oDisplay("item") = "Citizen Acct" Then
				response.write "<td align=""center""><a href=""../purchases/viewjournal.asp?uid=" & oDisplay("userid") & "&pid=" & oDisplay("paymentid") & "&rt=c&it=ci&jet=d"">" & oDisplay("paymentid") & "</a></td>"
			Else 
				response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("paymentid") & """>" & oDisplay("paymentid") & "</a></td>"
			End If 
			response.write "<td align=""center"">" & oDisplay("paymentdate") & "</td>"
			response.write "<td align=""center"" valign=""top"">" & oDisplay("userfname") & " " & oDisplay("userlname") & "<br />" & FormatPhoneNumber(oDisplay("userhomephone")) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("checkamt"), 2) & "<br />" & oDisplay("checkno") & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cashamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cardamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cccsubtotal"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("otheramt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("memoamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("paymenttotal"),2) & "</td>"
			dCheckTotal = dCheckTotal + CDbl(oDisplay("checkamt"))
			dCashTotal = dCashTotal + CDbl(oDisplay("cashamt"))
			dCardTotal = dCardTotal + CDbl(oDisplay("cardamt"))
			dOtherTotal = dOtherTotal + CDbl(oDisplay("otheramt"))
			dMemoTotal = dMemoTotal + CDbl(oDisplay("memoamt"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("paymenttotal"))
			dCCCTotal = dCCCTotal + CDbl(oDisplay("cccsubtotal"))
			response.write "</tr>"
			oDisplay.MoveNext
			response.flush
		Loop 
		' Totals Row
		If bgcolor="#eeeeee" Then
			bgcolor="#ffffff" 
		Else
			bgcolor="#eeeeee"
		End If	
		response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """ class=""totalrow""><td colspan=""3"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dCheckTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCashTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCardTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCCCTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dOtherTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dMemoTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"
	
	Else
		' No data found'
		response.write vbcrlf & "<div id=""nodatafound"">No data could be found that matched your search criteria.</div>"
	End If 

	oDisplay.Close
	Set oDisplay = Nothing 
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' boolean PullClassPaymentData( sFrom, sWhereClause, oDisplay )
'------------------------------------------------------------------------------------------------------------
Function PullClassPaymentData( ByVal sFrom, ByVal sWhereClause, ByRef oDisplay )
	Dim oRs, bHasData, sSql

	sSql = "SELECT paymentid, orgid, userid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone, paymenttotal, paymentdate, journalentrytype, amount, "
	sSql = sSql & " paymenttypename, checkno, isothermethod, requirescash, requirescreditcard, requirescitizenaccount, "
	sSql = sSql & " requirescheckno, paymentlocationname, adminlocationid, adminuserid, item, [Transaction ID] "
	sSql = sSql & " FROM " & sFrom & " " & sWhereClause
	sSql = sSql & " ORDER BY paymentid" 
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF then
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("paymentid")) <> CLng(iOldPaymentId) Then
				oDisplay.addnew 
				oDisplay("paymentid") = oRs("paymentid")
				oDisplay("paymentdate") = DateValue(oRs("paymentdate"))
				oDisplay("item") = oRs("item")
				oDisplay("userid") = oRs("userid")
				oDisplay("userfname") = oRs("userfname")
				oDisplay("userlname") = oRs("userlname")
				oDisplay("userhomephone") = oRs("userhomephone")
				oDisplay("paymenttotal") = oRs("paymenttotal")
				oDisplay("checkamt") = 0.00
				oDisplay("cashamt") = 0.00
				oDisplay("cardamt") = 0.00
				oDisplay("cccsubtotal") = 0.00
				oDisplay("otheramt") = 0.00
				oDisplay("memoamt") = 0.00
				dCCCSubTotal = 0.00
				iOldPaymentId = CLng(oRs("paymentid"))
			End If 
			If oRs("requirescheckno") Then
				oDisplay("checkamt") = oRs("amount")
				oDisplay("checkno") = oRs("checkno")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRs("amount"))
			End If 
			If oRs("requirescash") Then
				oDisplay("cashamt") = oRs("amount")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRs("amount"))
			End If 
			If oRs("requirescreditcard") Then
				oDisplay("cardamt") = oRs("amount")
				dCCCSubTotal = dCCCSubTotal + CDbl(oRs("amount"))
			End If 
			If oRs("isothermethod") Then
				oDisplay("otheramt") = oRs("amount")
			End If 
			If oRs("requirescitizenaccount") Then
				oDisplay("memoamt") = oRs("amount")
			End If 
			oDisplay("cccsubtotal") = dCCCSubTotal

			oDisplay.Update
			oRs.MoveNext
		Loop
	Else
		bHasData = False
	End If 
	
	oRs.Close
	Set oRs = Nothing

	PullClassPaymentData = bHasData

End Function 


'------------------------------------------------------------------------------------------------------------
' Display_Class_Refund_Report sWhereClause 
'------------------------------------------------------------------------------------------------------------
Sub Display_Class_Refund_Report( ByVal sWhereClause )
	Dim sSql, oRequests, oDisplay, iOldPaymentId, dVoucherTotal,  dCardTotal, dMemoTotal, dGrandTotal, dSubTotal

	iOldPaymentId = CLng(0) 

	sSql = "SELECT paymentid, orgid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
	sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone, paymentdate, amount, isccrefund, priorbalance "
	sSql = sSql & " FROM egov_class_to_refund_method " & sWhereClause
	sSql = sSql & " ORDER BY paymentid" 
	'response.write sSql & "<br><br>"

	' for export to CSV
	'session("DISPLAYQUERY") = sSql

	Set oRequests = Server.CreateObject("ADODB.Recordset")
	oRequests.Open sSQL, Application("DSN"), 3, 1

	If oRequests.EOF then
		' EMPTY
		response.write "<div id=""nodatafound"">No data could be found that matched your search criteria.</div>"
	Else
		' Got some data now make a holding recordset
		Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
		oDisplay.fields.append "paymentid", adInteger, , adFldUpdatable
		oDisplay.fields.append "paymentdate", adVariant, 10, adFldUpdatable
		oDisplay.fields.append "userfname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "userlname", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "userhomephone", adVarChar, 50, adFldUpdatable
		oDisplay.fields.append "voucheramt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "cardamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "subtotal", adCurrency, , adFldUpdatable
		oDisplay.fields.append "memoamt", adCurrency, , adFldUpdatable
		oDisplay.fields.append "total", adCurrency, , adFldUpdatable

		oDisplay.CursorLocation = 3
		'oDisplay.CursorType = 3

		oDisplay.open 

		' Loop through and build the display recordset.
		Do While Not oRequests.EOF
			If CLng(oRequests("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("paymentid") = oRequests("paymentid")
				oDisplay("paymentdate") = DateValue(oRequests("paymentdate"))
				If Not IsNull(oRequests("userfname")) Then 
					oDisplay("userfname") = oRequests("userfname")
				End If 
				If Not IsNull(oRequests("userlname")) Then 
					oDisplay("userlname") = oRequests("userlname")
				End If 
				oDisplay("userhomephone") = oRequests("userhomephone")
				oDisplay("voucheramt") = 0.00
				oDisplay("cardamt") = 0.00
				oDisplay("subtotal") = 0.00
				oDisplay("memoamt") = 0.00
				oDisplay("total") = 0.00
				iOldPaymentId = CLng(oRequests("paymentid"))
			End If 
			If oRequests("isccrefund") Then
				' Credit Card Refund
				oDisplay("cardamt") = oRequests("amount")
				oDisplay("subtotal") = CDbl(oDisplay("subtotal")) + CDbl(oRequests("amount"))
			Else 
				If IsNull(oRequests("priorbalance")) Then
					' Voucher Issued
					oDisplay("voucheramt") = oRequests("amount")
					oDisplay("subtotal") = CDbl(oDisplay("subtotal")) + CDbl(oRequests("amount"))
				Else
					' Refund To Memo account
					oDisplay("memoamt") = oRequests("amount")
				End If 
			End If 
			oDisplay("total") = CDbl(oDisplay("total")) + CDbl(oRequests("amount"))

			oDisplay.Update
			oRequests.MoveNext
		Loop

		' Show results
		oDisplay.MoveFirst

		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2""  border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist""><th>Receipt</th><th>Date</th><th>Payee</th>"
		response.write "<th>Voucher<br />Amount</th><th>Card Amt</th><th>Total Card<br />&amp; Voucher</th><th>Memo Amt</th><th>Total<br />Refund</th></tr>"
		bgcolor = "#eeeeee"
		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then
				bgcolor="#ffffff" 
			Else
				bgcolor="#eeeeee"
			End If			
			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"
			response.write "<td align=""center""><a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("paymentid") & """>" & oDisplay("paymentid") & "</a></td>"
			response.write "<td align=""center"">" & oDisplay("paymentdate") & "</td>"
			response.write "<td align=""center"" valign=""top"">" & oDisplay("userfname") & " " & oDisplay("userlname") & "<br />" & FormatPhoneNumber(oDisplay("userhomephone")) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("voucheramt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("cardamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("subtotal"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("memoamt"), 2) & "</td>"
			response.write "<td align=""right"">" & FormatNumber(oDisplay("total"),2) & "</td>"
			dVoucherTotal = dVoucherTotal + CDbl(oDisplay("voucheramt"))
			dCardTotal = dCardTotal + CDbl(oDisplay("cardamt"))
			dSubTotal = dSubTotal + CDbl(oDisplay("subtotal"))
			dMemoTotal = dMemoTotal + CDbl(oDisplay("memoamt"))
			dGrandTotal = dGrandTotal + CDbl(oDisplay("total"))
			response.write "</tr>"
			oDisplay.MoveNext
		Loop 
		' Totals Row
		If bgcolor="#eeeeee" Then
			bgcolor="#ffffff" 
		Else
			bgcolor="#eeeeee"
		End If	
		response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """ class=""totalrow""><td colspan=""3"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dVoucherTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dCardTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dSubTotal, 2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dMemoTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dGrandTotal,2) & "</td>"
		response.write "</tr>"

		response.write vbcrlf & "</table>"

		oDisplay.Close
		Set oDisplay = Nothing 

	End If 

	oRequests.Close
	Set oRequests = Nothing 

End Sub 


'------------------------------------------------------------------------------
' Display_Class_Acct_Dist_Details sWhereClause
'------------------------------------------------------------------------------
Sub Display_Class_Acct_Dist_Details( ByVal sWhereClause )
	Dim sSql, oRs, oDisplay, iOldAccountId, iOldPaymentId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal, bHasData

	iOldAccountId = CLng(0) 
	iOldPaymentId = CLng(0)
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)
	bHasData = False 

	' Got some data now make a holding recordset
	Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
	oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
	oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oDisplay.fields.append "receiptno", adInteger, , adFldUpdatable
	oDisplay.fields.append "paymentdate", adDBTimeStamp, , adFldUpdatable
	oDisplay.fields.append "paymenttypeid", adInteger, , adFldUpdatable
	oDisplay.fields.append "journalentrytypeid", adInteger, , adFldUpdatable
	oDisplay.fields.append "userid", adInteger, , adFldUpdatable
	oDisplay.fields.append "creditamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "debitamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "totalamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oDisplay.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable

	oDisplay.CursorLocation = 3
	'oDisplay.CursorType = 3

	oDisplay.Open 

	sSql = "SELECT accountname, accountnumber, accountid, entrytype, paymentid, amount, paymentdate, "
	sSql = sSql & "paymenttypeid, userid, journalentrytypeid, ispaymentaccount, iscitizenaccount "
	sSql = sSql & "FROM egov_class_to_account_dist_method "
	sSql = sSql & sWhereClause 
	sSql = sSql & " ORDER BY accountid, paymentid, entrytype"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		bHasData = True 
		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) <> iOldAccountId Or CLng(oRs("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("accountid")        = oRs("accountid")
				oDisplay("accountname")      = oRs("accountname")
				oDisplay("accountnumber")    = oRs("accountnumber")
				oDisplay("ispaymentaccount") = oRs("ispaymentaccount")

				If oRs("accountname") = "Citizen Accounts" Then 
  					oDisplay("iscitizenaccount") = True 
				Else  
		  			oDisplay("iscitizenaccount") = False 
				End If 

				oDisplay("receiptno")          = oRs("paymentid")
				oDisplay("paymentdate")        = oRs("paymentdate")
				oDisplay("paymenttypeid")      = oRs("paymenttypeid")
				oDisplay("journalentrytypeid") = oRs("journalentrytypeid")
				'Response.write "[" & oRs("userid") & "]<br />"
				oDisplay("userid")             = oRs("userid")
				oDisplay("creditamt")          = CDbl(0.00)
				oDisplay("debitamt")           = CDbl(0.00)
				oDisplay("totalamt")           = CDbl(0.00)
				iOldAccountId                  = CLng(oRs("accountid"))
				iOldPaymentId                  = CLng(oRs("paymentid"))
			End If 

			If oRs("entrytype") = "credit" Then
  				oDisplay("creditamt") = oDisplay("creditamt") + CDbl(oRs("amount"))
		  		oDisplay("totalamt")  = oDisplay("totalamt")  + CDbl(oRs("amount"))
			End If 

			If oRs("entrytype") = "debit" Then
  				oDisplay("debitamt") = oDisplay("debitamt") - CDbl(oRs("amount"))
		  		oDisplay("totalamt") = oDisplay("totalamt") - CDbl(oRs("amount"))
			End If 
			oDisplay.Update
			oRs.MoveNext
		Loop
	End If 
	oRs.Close 
	Set oRs = Nothing 

	'Citizen Accounts
	sSql = "SELECT accountname, accountnumber, accountid, entrytype, paymentid, amount, paymentdate, "
	sSql = sSql & "paymenttypeid, userid, journalentrytypeid, ispaymentaccount, iscitizenaccount "
	sSql = sSql & "FROM egov_class_to_citizen_acct_dist "
	sSql = sSql & sWhereClause  
	sSql = sSql & " ORDER BY accountid, paymentid, entrytype"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		bHasData = True 
		iOldAccountId = CLng(0)

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) <> iOldAccountId Or CLng(oRs("paymentid")) <> iOldPaymentId Then
				oDisplay.addnew 
				oDisplay("accountid")          = oRs("accountid")
				oDisplay("accountname")        = oRs("accountname") 
				oDisplay("accountnumber")      = oRs("accountnumber")
				oDisplay("ispaymentaccount")   = True 
				oDisplay("iscitizenaccount")   = True 
				oDisplay("receiptno")          = oRs("paymentid")
				oDisplay("paymentdate")        = oRs("paymentdate")
				oDisplay("paymenttypeid")      = oRs("paymenttypeid")
				oDisplay("journalentrytypeid") = oRs("journalentrytypeid")
				oDisplay("userid")             = oRs("userid")
				oDisplay("creditamt")          = CDbl(0.00)
				oDisplay("debitamt")           = CDbl(0.00)
				oDisplay("totalamt")           = CDbl(0.00)
				iOldAccountId                  = CLng(oRs("accountid"))
				iOldPaymentId                  = CLng(oRs("paymentid"))
			End If 

			If oRs("entrytype") = "credit" Then
  				oDisplay("creditamt") = oDisplay("creditamt") + CDbl(oRs("amount"))
		   	oDisplay("totalamt")  = oDisplay("totalamt")  + CDbl(oRs("amount"))
			End If 

			If oRs("entrytype") = "debit" Then
  				oDisplay("debitamt") = oDisplay("debitamt") - CDbl(oRs("amount"))
		  		oDisplay("totalamt") = oDisplay("totalamt") - CDbl(oRs("amount"))
			End If 
			oDisplay.Update
			oRs.MoveNext
		Loop
		 
	End If 
	oRs.Close
	Set oRs = Nothing 


	If bHasData Then 
		'sort the Display recordset
		oDisplay.Sort = "ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC, receiptno ASC"

		' Show results
		oDisplay.MoveFirst
		'response.write "<div class=""receiptpaymentshadow"">" & vbcrlf
		response.Write "<table cellspacing=""0"" cellpadding=""2"" border=""0"" width=""100%"" class=""receiptpayment"">" & vbcrlf
		response.write "  <tr class=""tablelist"">" & vbcrlf
		response.write "      <th>Account Name</th>" & vbcrlf
		response.write "      <th>Account Number</th>" & vbcrlf
		response.write "      <th>Receipt No.</th>" & vbcrlf
		response.write "      <th>Date</th>" & vbcrlf
		response.write "      <th>Total Amt<br />Credited</th>" & vbcrlf
		response.write "      <th>Total Amt<br />Debited</th>" & vbcrlf
		response.write "      <th>Total Amt<br />Transfered</th>" & vbcrlf
		response.write "  </tr>" & vbcrlf

		bgcolor         = "#eeeeee"
		iOldAccountId   = CLng(0)
		dCreditSubTotal = CDbl(0.00)
		dDebitSubTotal  = CDbl(0.00)
		dSubTotal       = CDbl(0.00)

		Do While Not oDisplay.EOF
			If bgcolor="#eeeeee" Then 
				bgcolor="#ffffff" 
			Else 
				bgcolor="#eeeeee"
			End If 

  			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>"

		  	If iOldAccountId <> CLng(oDisplay("accountid")) Then 
				   'Put out a sub total row
    				If iOldAccountId <> CLng(0) Then 
						response.write vbcrlf & "<tr class=""totalrow"">"
						response.write "<td colspan=""4"" align=""right"">Sub-Total:</td>"
						response.write "<td align=""right"">" & FormatNumber(dCreditSubTotal, 2) & "</td>"
						response.write "<td align=""right"">" & FormatNumber(dDebitSubTotal, 2)  & "</td>"
						response.write "<td align=""right"">" & FormatNumber(dSubTotal,2)        & "</td>" 
						response.write "</tr>"
  		  			End If 

  		  			response.write "<td align=""left"">"   & oDisplay("accountname")   & "</td>" 
    				response.write "<td align=""center"">" & oDisplay("accountnumber") & "</td>" 

					iOldAccountId   = CLng(oDisplay("accountid"))
					dCreditSubTotal = CDbl(0.00)
					dDebitSubTotal  = CDbl(0.00)
					dSubTotal       = CDbl(0.00)
		  	Else 
				'Need place holders 
				response.write "<td>&nbsp;</td>" 
				response.write "<td>&nbsp;</td>" 
		  	End If 

		  	If clng(oDisplay("journalentrytypeid")) > clng(2) Then 
				'citizen account activity
				response.write "<td align=""center"">"
				response.write "<a href=""../purchases/viewjournal.asp?uid=" & oDisplay("userid") & "&pid=" & oDisplay("receiptno") & "&rt=c&it=ci&jet=d"">" & oDisplay("receiptno") & "</a>"
				response.write "</td>" 
  			Else 
				'purchase
				response.write "<td align=""center"">"
				response.write "<a href=""../classes/view_receipt.asp?iPaymentId=" & oDisplay("receiptno") & """>" & oDisplay("receiptno") & "</a>"
				response.write "</td>" 
  			End If 

		  	response.write "<td align=""right"">" & FormatDateTime(oDisplay("paymentdate"), 2) & "</td>" 
  			response.write "<td align=""right"">" & FormatNumber(oDisplay("creditamt"), 2)     & "</td>" 
		  	response.write "<td align=""right"">" & FormatNumber(oDisplay("debitamt"), 2)      & "</td>" 
  			response.write "<td align=""right"">" & FormatNumber(oDisplay("totalamt"), 2)      & "</td>" 

  			dCreditSubTotal = dCreditSubTotal + CDbl(oDisplay("creditamt"))
		  	dTotalCredit    = dTotalCredit + CDbl(oDisplay("creditamt"))
  			dDebitSubTotal  = dDebitSubTotal + CDbl(oDisplay("debitamt"))
		  	dTotalDebit     = dTotalDebit + CDbl(oDisplay("debitamt"))
  			dSubTotal       = dSubTotal + CDbl(oDisplay("totalamt"))
		  	dGrandTotal     = dGrandTotal + CDbl(oDisplay("totalamt"))

  			response.write "  </tr>" & vbcrlf

			oDisplay.MoveNext
		Loop 

		'Put out a sub total row
		If iOldAccountId <> CLng(0) Then 
			response.write vbcrlf & "<tr class=""totalrow"">"
			response.write "<td colspan=""4"" align=""right"">Sub-Total:</td>" & vbcrlf
			response.write "<td align=""right"">" & FormatNumber(dCreditSubTotal, 2) & "</td>" 
			response.write "<td align=""right"">" & FormatNumber(dDebitSubTotal, 2)  & "</td>" 
			response.write "<td align=""right"">" & FormatNumber(dSubTotal,2)        & "</td>"
			response.write "</tr>"
		End If 

		'Totals Row
		response.write vbcrlf & "<tr class=""totalrow"">" 
		response.write "<td colspan=""4"" align=""right"">Totals:</td>" 
		response.write "<td align=""right"">" & FormatNumber( dTotalCredit, 2 ) & "</td>"
		response.write "<td align=""right"">" & FormatNumber( dTotalDebit, 2 )  & "</td>" 
		response.write "<td align=""right"">" & FormatNumber( dGrandTotal, 2 )  & "</td>" 
		response.write "</tr>" 
		response.write vbcrlf & "</table>" 
		'response.write vbcrlf & "</div>"
	
	Else
		response.write "<div id=""nodatafound"">No data could be found that matched your search criteria.</div>"
	End If

	oDisplay.Close
	Set oDisplay = Nothing

End Sub 


'------------------------------------------------------------------------------
' Display_Class_Acct_Dist_Summary sWhereClause
'------------------------------------------------------------------------------
Sub Display_Class_Acct_Dist_Summary( ByVal sWhereClause )
	Dim sSql, oRs, oDisplay, iOldAccountId, dTotal, dTotalCredit, dTotalDebit, dGrandTotal, bHasData

	iOldAccountId = CLng(0) 
	dTotal = CDbl(0.00)
	dTotalCredit = CDbl(0.00)
	dTotalDebit = CDbl(0.00)
	dGrandTotal = CDbl(0.00)
	bHasData = False 

	' Got some data now make a holding recordset
	Set oDisplay = server.CreateObject("ADODB.RECORDSET") 
	oDisplay.fields.append "accountid", adInteger, , adFldUpdatable
	oDisplay.fields.append "accountname", adVarChar, 50, adFldUpdatable
	oDisplay.fields.append "accountnumber", adVarChar, 20, adFldUpdatable
	oDisplay.fields.append "creditamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "debitamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "totalamt", adCurrency, , adFldUpdatable
	oDisplay.fields.append "ispaymentaccount", adBoolean, , adFldUpdatable
	oDisplay.fields.append "iscitizenaccount", adBoolean, , adFldUpdatable

	oDisplay.CursorLocation = 3
	'oDisplay.CursorType = 3

	oDisplay.open 

	' Pull all account data except the citizen accounts
	sSql = "SELECT accountname, accountnumber, accountid, entrytype, ispaymentaccount, iscitizenaccount, "
	sSql = sSql & "SUM(amount) AS amount "
	sSql = sSql & "FROM egov_class_to_account_dist_method "
	sSql = sSql & sWhereClause
	sSql = sSql & " GROUP BY accountname, accountnumber, accountid, entrytype, ispaymentaccount, iscitizenaccount "
	sSql = sSql & " ORDER BY accountid, entrytype"
	'response.write "<!-- " & sSql & " --><br /><br />"
	'response.write  sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) <> iOldAccountId Then
				oDisplay.addnew 
				oDisplay("accountid")        = oRs("accountid")
				oDisplay("accountname")      = oRs("accountname") 
				oDisplay("accountnumber")    = oRs("accountnumber")
				oDisplay("ispaymentaccount") = oRs("ispaymentaccount")
				oDisplay("iscitizenaccount") = oRs("iscitizenaccount") 
''				If sRptType = "Detail" Then
 '' 					oDisplay("paymentid") = oRs("paymentid")
''				End If 
				oDisplay("creditamt") = 0.00
				oDisplay("debitamt") = 0.00
				oDisplay("totalamt") = 0.00
				iOldAccountId = CLng(oRs("accountid"))
			End If 
			If oRs("entrytype") = "credit" Then
				oDisplay("creditamt") = CDbl(oRs("amount"))
				'dTotal = CDbl(oRs("amount"))
				oDisplay("totalamt") = CDbl(oDisplay("totalamt")) + CDbl(oRs("amount"))
			End If 
			If oRs("entrytype") = "debit" Then
				oDisplay("debitamt") = -CDbl(oRs("amount"))
				'dTotal = -CDbl(oRs("amount"))
				oDisplay("totalamt") = CDbl(oDisplay("totalamt")) - CDbl(oRs("amount"))
			End If 
				
			oDisplay.Update
			oRs.MoveNext
		Loop
	End If 

	oRs.Close
	Set oRs = Nothing 

	'Get the citizen accounts summary here
	sSql = "SELECT accountname, accountnumber, accountid, entrytype, ispaymentaccount, iscitizenaccount, "
	sSql = sSql & "SUM(amount) AS amount "
	sSql = sSql & "FROM egov_class_to_citizen_acct_dist "
	sSql = sSql & sWhereClause 
	sSql = sSql & " GROUP BY accountname, accountnumber, accountid, entrytype, ispaymentaccount, iscitizenaccount "
	sSql = sSql & " ORDER BY accountid, entrytype"
	'response.write sSql & "<br />"
	'response.end

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF then
		bHasData = True 

		' Loop through and build the display recordset.
		Do While Not oRs.EOF
			If CLng(oRs("accountid")) <> iOldAccountId Then
				oDisplay.addnew 
				oDisplay("accountid") = oRs("accountid")
				oDisplay("accountname") = oRs("accountname") 
				oDisplay("accountnumber") = oRs("accountnumber")
				'oDisplay("ispaymentaccount") = oRs("ispaymentaccount")
				'oDisplay("iscitizenaccount") = oRs("iscitizenaccount")
				oDisplay("ispaymentaccount") = True 
				oDisplay("iscitizenaccount") = True 
''				If sRptType = "Detail" Then
''					oDisplay("paymentid") = oRs("paymentid")
''				End If 
				oDisplay("creditamt") = 0.00
				oDisplay("debitamt") = 0.00
				oDisplay("totalamt") = 0.00
				iOldAccountId = CLng(oRs("accountid"))
			End If 
			If oRs("entrytype") = "credit" Then
				oDisplay("creditamt") = oDisplay("creditamt") + CDbl(oRs("amount"))
				'dTotal = CDbl(oRs("amount"))
				oDisplay("totalamt") = CDbl(oDisplay("totalamt")) + CDbl(oRs("amount"))
			End If 
			If oRs("entrytype") = "debit" Then
				oDisplay("debitamt") = oDisplay("debitamt") - CDbl(oRs("amount"))
				'dTotal = -CDbl(oRs("amount"))
				oDisplay("totalamt") = CDbl(oDisplay("totalamt")) - CDbl(oRs("amount"))
			End If 
				
			oDisplay.Update
			oRs.MoveNext
		Loop
	End If 

	oRs.Close
	Set oRs = Nothing 
		
	If bHasData Then 
		'sort the Display recordset
		oDisplay.Sort = "ispaymentaccount DESC, iscitizenaccount ASC, accountname ASC, accountnumber ASC "

		' Show results
		oDisplay.MoveFirst

		response.Write vbcrlf & "<table cellspacing=""0"" cellpadding=""2"" border=""0"" width=""100%"" class=""receiptpayment"">"
		response.write vbcrlf & "<tr class=""tablelist"">" 
		response.write "<th>Account Name</th><th>Account Number</th><th>Total Amt<br />Credited</th>" 
		response.write "<th>Total Amt<br />Debited</th><th>Total Amt<br />Transfered</th></tr>" 

		bgcolor = "#eeeeee"
		Do While Not oDisplay.EOF
			bgcolor = changeBGColor(bgcolor,"#eeeeee","#ffffff")

			response.write vbcrlf & "<tr bgcolor=""" &  bgcolor  & """>" & vbcrlf
			response.write "<td align=""left"">"   & oDisplay("accountname")                & "</td>"
			response.write "<td align=""center"">" & oDisplay("accountnumber")              & "</td>"
			response.write "<td align=""right"">"  & FormatNumber(oDisplay("creditamt"), 2) & "</td>"
			response.write "<td align=""right"">"  & FormatNumber(oDisplay("debitamt"), 2)  & "</td>" 
			response.write "<td align=""right"">"  & FormatNumber(oDisplay("totalamt"), 2)  & "</td>" 

			dTotalCredit = dTotalCredit + CDbl(oDisplay("creditamt"))
			dTotalDebit  = dTotalDebit  + CDbl(oDisplay("debitamt"))
			dGrandTotal  = dGrandTotal  + dTotalCredit - dTotalDebit

			response.write "</tr>" 
			oDisplay.MoveNext
		Loop 

		'Totals Row
		response.write vbcrlf & "<tr class=""totalrow"">" 
		response.write "<td colspan=""2"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalCredit, 2)                & "</td>"
		response.write "<td align=""right"">" & FormatNumber(dTotalDebit, 2)                 & "</td>"
		response.write "<td align=""right"">" & FormatNumber((dTotalCredit + dTotalDebit),2) & "</td>" 
		response.write "</tr>"

		response.write vbcrlf & "</table>" 

	Else
		response.write "<div id=""nodatafound"">No data could be found that matched your search criteria.</div>"
	End If 


	oDisplay.Close
	Set oDisplay = Nothing 

End Sub 


%>
