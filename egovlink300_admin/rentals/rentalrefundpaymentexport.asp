<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalrefundpayment.asp
' AUTHOR: SteveLoar
' CREATED: 12/17/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This report has rental refunds by payment type - Part of Menlo Park Project
'
' MODIFICATION HISTORY
' 1.0   12/17/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, oSchema, iOldPaymentId, dCashTotal, dCheckTotal, dCardtotal, dMemoTotal, dGrandTotal
Dim iLocationId, iAdminUserId, toDate, fromDate, sRptTitle, iPaymentLocationId, dOtherTotal, dCCCTotal
Dim dCCCSubTotal, today, sWhereClause
Dim from_time, to_time, where_time

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Rental_Refund_Payment_" & sDate & ".xls"

sRptTitle = "<tr><th></th><th>Refund Payments</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

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
	'fromDate = cdate(Month(today)& "/1/" & Year(today)) 
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

If request("adminuserid") = "0" Then
	iAdminUserId = 0
Else
	iAdminUserId = CLng(request("adminuserid"))
End If 

' BUILD SQL WHERE CLAUSE
sWhereClause = " WHERE orgid = " & session("orgid") 

'sWhereClause = sWhereClause & " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
'sRptTitle = sRptTitle & "<tr><th>Refund Date >= " & fromDate & "</th><th>AND Refund Date < " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
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
	sRptTitle = sRptTitle & "<th>AND Transaction Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th><th></th></tr>"
Else 
	where_time = CDate( todate & " " & to_time )
	sWhereClause = sWhereClause & " AND paymentDate <= '" & where_time & "' "
	sRptTitle = sRptTitle & "<th>AND Transaction Date <= " & where_time & "</th><th></th><th></th><th></th><th></th></tr>"
End If 

If iLocationId > 0 Then
	sWhereClause = sWhereClause & " AND adminlocationid = " & iLocationId
	sRptTitle = sRptTitle & "<tr><th>Admin Location: " & GetLocationName( iLocationId )  & "</th><th></th><th></th><th></th><th></th><th><th></th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Admin Location: All Locations</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If iAdminUserId > 0 Then
	sWhereClause = sWhereClause & " AND adminuserid = " & iAdminUserId
	sRptTitle = sRptTitle & "<tr><th>Admin: " & GetAdminName( iAdminUserId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Admin: All</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

iOldPaymentId = CLng(0) 

' Make a holding recordset
Set oSchema = server.CreateObject("ADODB.RECORDSET") 
oSchema.fields.append "Receipt Number", adInteger, , adFldUpdatable
oSchema.fields.append "Transaction Date", adVariant, 10, adFldUpdatable
oSchema.fields.append "Transaction Time", adVarChar, 20, adFldUpdatable
oSchema.fields.append "First Name", adVarChar, 50, adFldUpdatable
oSchema.fields.append "Last Name", adVarChar, 50, adFldUpdatable
oSchema.fields.append "Phone", adVarChar, 50, adFldUpdatable
oSchema.fields.append "Voucher", adCurrency, , adFldUpdatable
oSchema.fields.append "Card", adCurrency, , adFldUpdatable
oSchema.fields.append "Subtotal", adCurrency, , adFldUpdatable
oSchema.fields.append "Memo", adCurrency, , adFldUpdatable
oSchema.fields.append "Total", adCurrency, , adFldUpdatable

oSchema.CursorLocation = 3

oSchema.open 

sSql = "SELECT paymentid, orgid, ISNULL(userfname,'') AS userfname, ISNULL(userlname,'') AS userlname, "
sSql = sSql & " ISNULL(userhomephone,'') AS userhomephone, paymentdate, amount, isccrefund, priorbalance "
sSql = sSql & " FROM egov_rentals_to_refund_method " & sWhereClause
sSql = sSql & " ORDER BY paymentid" 
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then
	dVoucherTotal = CDbl(0.00)
	dCardTotal = CDbl(0.00)
	dSubTotal = CDbl(0.00)
	dMemoTotal = CDbl(0.00)
	dGrandTotal = CDbl(0.00)

	' Loop through and build the display recordset.
	Do While Not oRs.EOF
		If CLng(oRs("paymentid")) <> iOldPaymentId Then
			iOldPaymentId = CLng(oRs("paymentid"))

			oSchema.addnew 
			oSchema("Receipt Number") = oRs("paymentid")
			oSchema("Transaction Date") = DateValue(oRs("paymentdate"))
			oSchema("Transaction Time") = FormatDateTime(oRs("paymentdate"),3)
			oSchema("First Name") = oRs("userfname")
			oSchema("Last Name") = oRs("userlname")
			oSchema("Phone") = FormatPhoneNumber(oRs("userhomephone"))
			oSchema("Voucher") = 0.00
			oSchema("Card") = 0.00
			oSchema("Subtotal") = 0.00
			oSchema("Memo") = 0.00
			oSchema("Total") = 0.00
		End If 
		If oRs("isccrefund") Then
			' Credit Card Refund
			oSchema("Card") = oRs("amount")
			oSchema("Subtotal") = CDbl(oSchema("Subtotal")) + CDbl(oRs("amount"))
			dCardTotal = dCardTotal + CDbl(oRs("amount"))
			dSubTotal = dSubTotal + CDbl(oRs("amount"))
		Else 
			If IsNull(oRs("priorbalance")) Then
				' Voucher Issued
				oSchema("Voucher") = oRs("amount")
				oSchema("Subtotal") = CDbl(oSchema("Subtotal")) + CDbl(oRs("amount"))
				dVoucherTotal = dVoucherTotal + CDbl(oRs("amount"))
				dSubTotal = dSubTotal + CDbl(oRs("amount"))
			Else
				' Refund To Memo account
				oSchema("Memo") = oRs("amount")
				dMemoTotal = dMemoTotal + CDbl(oRs("amount"))
			End If 
		End If 
		oSchema("Total") = CDbl(oSchema("Total")) + CDbl(oRs("amount"))
		dGrandTotal = dGrandTotal + CDbl(oRs("amount"))
		oSchema.Update
		oRs.MoveNext
	Loop
Else
	' A blank row
	oSchema.addnew 
	oSchema("Receipt Number") = 0
	oSchema("First Name") = " "
	oSchema("Last Name") = " "
	oSchema("Phone") = " "
	oSchema("Voucher") = 0.00
	oSchema("Card") = 0.00
	oSchema("Subtotal") = 0.00
	oSchema("Total") = 0.00
	oSchema("Memo") = 0.00
	oSchema.Update
End If 

oRs.Close
Set oRs = Nothing

oSchema.MoveFirst

' Total Row
sTotalRow = "<tr><td></td><td></td><td></td><td></td><td></td><td>Total</td><td class=""moneystyle"">" & FormatNumber(dVoucherTotal, 2) & "</td><td class=""moneystyle"">" & FormatNumber(dCardTotal, 2) & "</td><td class=""moneystyle"">" & FormatNumber(dSubTotal, 2) & "</td><td class=""moneystyle"">" & FormatNumber(dMemoTotal,2) & "</td><td class=""moneystyle"">" & FormatNumber(dGrandTotal, 2) & "</td></tr>"

CreateExcelDownload sRptTitle, sTotalRow

oSchema.Close
Set oSchema = Nothing 


'--------------------------------------------------------------------------------------------------
' string sPhone = FormatPhoneNumber( Number )
'--------------------------------------------------------------------------------------------------
Function FormatPhoneNumber( ByVal Number )
	If Len(Number) = 10 Then
		FormatPhoneNumber = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhoneNumber = Number
	End If
End Function


'--------------------------------------------------------------------------------------------------
' string sName = GetAdminName( iUserId )
'--------------------------------------------------------------------------------------------------
Function GetAdminName( ByVal iUserId )
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
' string sLocation = GetLocationName( iLocationid )
'--------------------------------------------------------------------------------------------------
Function GetLocationName_old( ByVal iLocationid )
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


%>

<!-- #include file="../export/include_excel_export.asp" //-->

<!-- #include file="../includes/adovbs.inc" -->

<!-- #include file="rentalscommonfunctions.asp" //-->

