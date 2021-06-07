<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: receipt_payment_export.asp
' AUTHOR: SteveLoar
' CREATED: 01/10/07
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   7/17/2007	Steve Loar - INITIAL VERSION
' 1.1	10/4/2007	Steve Loar - Adding payments to citizen accounts to the report
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRequests, oSchema, iOldPaymentId, dCashTotal, dCheckTotal, dCardtotal, dMemoTotal, dGrandTotal, bHasData
Dim iLocationId, iAdminUserId, toDate, fromDate, sRptTitle, iPaymentLocationId, dOtherTotal, dCCCTotal, dCCCSubTotal
Dim from_time, to_time, where_time

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Receipt_Payments_" & sDate & ".xls"

sRptTitle = "<tr><th></th><th>Receipt Payments</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

' PROCESS REPORT FILTER VALUES
' PROCESS DATE VALUES
fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()
bHasData = False 

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

If request("paymentlocationid") = "" Then
	iPaymentLocationId = 0
Else
	iPaymentLocationId = CLng(request("paymentlocationid"))
End If 

' BUILD SQL WHERE CLAUSE
varWhereClause = " WHERE orgid = " & session("orgid") 

'varWhereClause = " WHERE (paymentDate >= '" & fromDate & "' AND paymentDate <= '" & DateAdd("d",1,toDate) & "') "
'sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & fromDate & "</th><th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
If from_time = "none" Then 
	varWhereClause = varWhereClause & " AND paymentDate >= '" & fromDate & "' "
	sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & fromDate & "</th>"
Else
	where_time = CDate( fromdate & " " & from_time )
	varWhereClause = varWhereClause & " AND paymentDate >= '" & where_time & "' "
	sRptTitle = sRptTitle & "<tr><th>Payment Date >= " & where_time & "</th>"
End If 

If to_time = "none" Then 
	varWhereClause = varWhereClause & " AND paymentDate <= '" & DateAdd("d",1,toDate) & "' "
	sRptTitle = sRptTitle & "<th>AND Payment Date <= " & DateAdd("d",1,toDate) & "</th><th></th><th></th><th></th><th></th></tr>"
Else 
	where_time = CDate( todate & " " & to_time )
	varWhereClause = varWhereClause & " AND paymentDate <= '" & where_time & "' "
	sRptTitle = sRptTitle & "<th>AND Payment Date <= " & where_time & "</th><th></th><th></th><th></th><th></th></tr>"
End If 

If iLocationId > 0 Then
	varWhereClause = varWhereClause & " AND adminlocationid = " & iLocationId
	sRptTitle = sRptTitle & "<tr><th>Admin Location: " & GetLocationName( iLocationId )  & "</th><th></th><th></th><th></th><th></th><th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Admin Location: All Locations</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 
If iAdminUserId > 0 Then
	varWhereClause = varWhereClause & " AND adminuserid = " & iAdminUserId
	sRptTitle = sRptTitle & "<tr><th>Admin: " & GetAdminName( iAdminUserId )  & "</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
Else
	sRptTitle = sRptTitle & "<tr><th>Admin: All</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 

If iPaymentLocationId > 0 Then
	If iPaymentLocationId = CLng(2) Then
		varWhereClause = varWhereClause & " AND paymentlocationid = 3 " 
		sRptTitle = sRptTitle & "<tr><th>Payment Location: Web Site Only</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	Else
		varWhereClause = varWhereClause & " AND paymentlocationid < 3 " 
		sRptTitle = sRptTitle & "<tr><th>Payment Location: Office Only</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
	End If 
Else
	sRptTitle = sRptTitle & "<tr><th>Payment Location: All Locations</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
End If 


iOldPaymentId = CLng(0) 
dCashTotal = CDbl(0.00)
dCheckTotal = CDbl(0.00)
dCardTotal = CDbl(0.00)
dOtherTotal = CDbl(0.00)
dMemoTotal = CDbl(0.00)
dGrandTotal = CDbl(0.00)
dCCCTotal = CDbl(0.00)
dCCCSubTotal = CDbl(0.00)


' Make a holding recordset
Set oSchema = server.CreateObject("ADODB.RECORDSET") 
oSchema.fields.append "receiptno", adInteger, , adFldUpdatable
oSchema.fields.append "paymentdate", adVariant, 10, adFldUpdatable
oSchema.fields.append "paymenttime", adVarChar, 20, adFldUpdatable
oSchema.fields.append "userfname", adVarChar, 50, adFldUpdatable
oSchema.fields.append "userlname", adVarChar, 50, adFldUpdatable
oSchema.fields.append "userhomephone", adVarChar, 50, adFldUpdatable
oSchema.fields.append "checkamt", adCurrency, , adFldUpdatable
oSchema.fields.append "checkno", adVarChar, 20, adFldUpdatable
oSchema.fields.append "cashamt", adCurrency, , adFldUpdatable
oSchema.fields.append "cardamt", adCurrency, , adFldUpdatable
oSchema.fields.append "cccsubtotal", adCurrency, , adFldUpdatable
oSchema.fields.append "otheramt", adCurrency, , adFldUpdatable
oSchema.fields.append "memoamt", adCurrency, , adFldUpdatable
oSchema.fields.append "paymenttotal", adCurrency, , adFldUpdatable

oSchema.CursorLocation = 3
'oSchema.CursorType = 3

oSchema.open 

' Get the class purchases
sSql = "SELECT paymentid, orgid, userid, ISNULL(userfname,'') as userfname, ISNULL(userlname,'') AS userlname, ISNULL(userhomephone,'') AS userhomephone, item, paymenttotal, paymentdate, journalentrytype, amount, paymenttypename, checkno, isothermethod, "
sSql = sSql & " requirescash, requirescreditcard, requirescitizenaccount, requirescheckno, paymentlocationname, adminlocationid, adminuserid, item, [Transaction ID] "
sSql = sSql & " FROM egov_class_to_payment_method " & varWhereClause
sSql = sSql & " ORDER BY paymentid" 
'Response.Write sSql & "<br /><br />"

Set oRequests = Server.CreateObject("ADODB.Recordset")
oRequests.Open sSQL, Application("DSN"), 3, 1

If Not oRequests.EOF Then
	bHasData = True 
	' Loop through and build the display recordset.
	Do While Not oRequests.EOF
		If CLng(oRequests("paymentid")) <> iOldPaymentId Then
			oSchema.addnew 
			oSchema("receiptno") = oRequests("paymentid")
			oSchema("paymentdate") = DateValue(oRequests("paymentdate"))
			oSchema("paymenttime") = FormatDateTime(oRequests("paymentdate"),3)
			oSchema("userfname") = oRequests("userfname")
			oSchema("userlname") = oRequests("userlname")
			oSchema("userhomephone") = FormatPhoneNumber(oRequests("userhomephone"))
			oSchema("paymenttotal") = oRequests("paymenttotal")
			oSchema("checkamt") = 0.00
			oSchema("cashamt") = 0.00
			oSchema("cardamt") = 0.00
			oSchema("cccsubtotal") = 0.00
			oSchema("otheramt") = 0.00
			oSchema("memoamt") = 0.00
			dCCCSubTotal = 0.00
			iOldPaymentId = CLng(oRequests("paymentid"))
		End If 
		If oRequests("requirescheckno") Then
			oSchema("checkamt") = oRequests("amount")
			dCheckTotal = dCheckTotal + CDbl(oRequests("amount"))
			dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
			oSchema("checkno") = oRequests("checkno")
			dCCCSubTotal = dCCCSubTotal + CDbl(oRequests("amount"))
			dCCCTotal = dCCCTotal + CDbl(oRequests("amount"))
		End If 
		If oRequests("requirescash") Then
			oSchema("cashamt") = oRequests("amount")
			dCashTotal = dCashTotal + CDbl(oRequests("amount"))
			dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
			dCCCSubTotal = dCCCSubTotal + CDbl(oRequests("amount"))
			dCCCTotal = dCCCTotal + CDbl(oRequests("amount"))
		End If 
		If oRequests("requirescreditcard") Then
			oSchema("cardamt") = oRequests("amount")
			dCardTotal = dCardTotal + CDbl(oRequests("amount"))
			dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
			dCCCSubTotal = dCCCSubTotal + CDbl(oRequests("amount"))
			dCCCTotal = dCCCTotal + CDbl(oRequests("amount"))
		End If 
		If oRequests("isothermethod") Then
			oSchema("otheramt") = oRequests("amount")
			dOtherTotal = dOtherTotal + CDbl(oRequests("amount"))
			dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
		End If
		If oRequests("requirescitizenaccount") Then
			oSchema("memoamt") = oRequests("amount")
			dMemoTotal = dMemoTotal + CDbl(oRequests("amount"))
			dGrandTotal = dGrandTotal + CDbl(oRequests("amount"))
		End If 
		oSchema("cccsubtotal") = dCCCSubTotal
			
		oSchema.Update
		oRequests.MoveNext
	Loop
	
End If 
oRequests.Close

If Not bHasData Then
	' Add A blank row
	oSchema.addnew 
	oSchema("receiptno") = 0
	oSchema("userfname") = " "
	oSchema("userlname") = " "
	oSchema("userhomephone") = " "
	oSchema("paymenttotal") = 0.00
	oSchema("checkamt") = 0.00
	oSchema("cashamt") = 0.00
	oSchema("cardamt") = 0.00
	oSchema("cccsubtotal") = 0.00
	oSchema("otheramt") = 0.00
	oSchema("memoamt") = 0.00
	oSchema.Update
End If 

oSchema.sort = "receiptno"
oSchema.MoveFirst

' Total Row
sTotalRow = "<tr><td></td><td></td><td></td><td></td><td></td><td>Total</td><td class=""moneystyle"">" & FormatNumber(dCheckTotal, 2) & "</td><td></td><td class=""moneystyle"">" & FormatNumber(dCashTotal, 2) & "</td><td class=""moneystyle"">" & FormatNumber(dCardTotal, 2) & "</td><td class=""moneystyle"">" & FormatNumber(dCCCTotal,2) & "</td><td class=""moneystyle"">" & FormatNumber(dOtherTotal, 2) & "</td><td class=""moneystyle"">" & FormatNumber(dMemoTotal, 2) & "</td><td class=""moneystyle"">" & FormatNumber(dGrandTotal,2) & "</td></tr>"

CreateExcelDownload sRptTitle, sTotalRow

oSchema.Close
Set oSchema = Nothing 


Set oRequests = Nothing



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
' string sNAme = GetAdminName( iUserId )
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


%>

<!-- #include file="../export/include_excel_export.asp" //-->

<!-- #include file="../includes/adovbs.inc" -->

