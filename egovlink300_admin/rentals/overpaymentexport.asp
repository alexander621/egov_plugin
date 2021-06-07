<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: overpaymentesport.asp
' AUTHOR: SteveLoar
' CREATED: 01/17/2013
' COPYRIGHT: Copyright 2013 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   7/17/2013	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, oSchema, toDate, fromDate, sSearch, iRentalId, sRentalIdType, iActualId, iReservationTypeId
Dim sReservationTypeSelector, sRenterName, iReservedStatus, sStatusSearch, sReservationId, iSupervisorUserId
Dim sPaymentId, sFrom, sReportTitle, sDepositsWhere, sStartDateTime, sEndBillingDateTime, sDatesSearch

sDepositsWhere = ""
sSearch = ""
sStartDateTime = ""
sEndBillingDateTime = ""
sDatesSearch = ""

sReportTitle = "<tr><td></td><td>Overpayment Report</td></tr>"

server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Rental_overpayments.xls"

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

sSearch = sSearch & " AND D.reservationstarttime BETWEEN '" & CDate(fromDate) & " 00:00 AM' AND '" & DateAdd("d",1,CDate(toDate)) &" 00:00 AM' "
sDatesSearch = sDatesSearch & " AND D.reservationstarttime BETWEEN '" & CDate(fromDate) & " 00:00 AM' AND '" & DateAdd("d",1,CDate(toDate)) &" 00:00 AM' "
sReportTitle = sReportTitle & "<tr><td>Reservation Start Time:</td><td>Between " & DateValue(CDate(fromDate)) & " AND " & DateValue(CDate(toDate)) & "</td></tr>"

If request("rentalid") <> "" Then
	iRentalId = request("rentalid")
	If iRentalId <> "0" Then 
		sRentalIdType = Left(iRentalId, 1)
		iActualId = Mid(iRentalId, 2)
		If sRentalIdType = "R" Then
			sSearch = sSearch & " AND R.rentalid = " & iActualId
			sDepositsWhere = sDepositsWhere & " AND R.rentalid = " & iActualId
			sDatesSearch = sDatesSearch & " AND D.rentalid = " & iActualId
			sReportTitle = sReportTitle & "<tr><td>Rental:</td><td>" & GetRentalName( iActualId ) & "</td></tr>"
		Else
			sSearch = sSearch & " AND L.locationid = " & iActualId
			sReportTitle = sReportTitle & "<tr><td>Rental:</td><td>" & GetLocationName( iActualId ) & " All Rentals</td></tr>"
		End If 
	End If 
Else
	iRentalId = "0"	  ' The all selection
	sReportTitle = sReportTitle & "<tr><td>Rental:</td><td>All Locations All Rentals</td></tr>"
End If 

If request("reservationid") <> "" Then
	sReservationId = request("reservationid")
	sSearch = sSearch & " AND D.reservationid = " & sReservationId
	sDepositsWhere = sDepositsWhere & " AND F.reservationid = " & sReservationId
	sReportTitle = sReportTitle & "<tr><td>Reservation Id: </td><td>" & sReservationId & "</td></tr>"
Else
	sReservationId = ""
End If 

sSql = "SELECT 'Rental Fees' AS overpayment, L.name, R.rentalname, F.reservationid, D.reservationstarttime, D.billingendtime, F.amount AS rate, F.duration, F.feeamount, F.paidamount, F.refundamount "
sSql = sSql & " FROM egov_rentalreservationdatefees F, egov_rentals R, egov_rentalreservationdates D, egov_class_location L "
sSql = sSql & " WHERE F.rentalid = R.rentalid AND R.orgid = "  & session("orgid")
sSql = sSql & " AND F.reservationdateid = D.reservationdateid AND R.locationid = L.locationid "
sSql = sSql & sSearch 
sSql = sSql & " AND F.paidamount > (F.feeamount + F.refundamount) AND F.paidamount > 0 "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 'Rental Items' AS overpayment, L.name, R.rentalname, F.reservationid, D.reservationstarttime, D.billingendtime, F.amount AS rate, quantity AS duration, F.feeamount, F.paidamount, F.refundamount "
sSql = sSql & " FROM egov_rentalreservationdateitems F, egov_rentals R, egov_rentalreservationdates D, egov_class_location L "
sSql = sSql & " WHERE D.rentalid = R.rentalid AND R.orgid = "  & session("orgid")
sSql = sSql & " AND F.reservationdateid = D.reservationdateid AND R.locationid = L.locationid "
sSql = sSql & sSearch 
sSql = sSql & " AND F.paidamount > (F.feeamount + F.refundamount) AND F.paidamount > 0 "
sSql = sSql & " UNION "
sSql = sSql & " SELECT 'Deposits and Charges' AS overpayment, L.name, R.rentalname, F.reservationid, NULL AS reservationstarttime, NULL AS billingendtime, F.amount AS rate, 0 AS duration, F.feeamount, F.paidamount, F.refundamount "
sSql = sSql & " FROM egov_rentalreservationfees F, egov_rentals R, egov_class_location L "
sSql = sSql & " WHERE F.rentalid = R.rentalid AND R.locationid = L.locationid AND R.orgid = "  & session("orgid")
sSql = sSql & sDepositsWhere 
sSql = sSql & " AND F.paidamount > (F.feeamount + F.refundamount) AND F.paidamount > 0 "
sSql = sSql & " AND F.reservationid IN ( SELECT reservationid FROM egov_rentalreservationdates D WHERE D.orgid = "  & session("orgid") 
sSql = sSql & sDatesSearch & " ) "
sSql = sSql & " ORDER BY 1, 2, 3, 5"

'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

response.write vbcrlf & "<html>"

response.write vbcrlf & "<style>  "
response.write " .moneystyle "
response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
response.write vbcrlf & "</style>"

response.write vbcrlf & "<body>"

response.write vbcrlf & "<table border=""1"">"
response.write sReportTitle

If Not oRs.EOF Then
	
	response.write "<tr><th>Overpayment Type</th><th>Location</th><th>Rental</th><th>Reservation Id</th><th>Reservation Date</th><th>Time</th><th>Base Rate</th><th>Duration/Quantity</th><th>Fee Amount</th><th>Paid Amount</th><th>Refunded Amount</th>"
	response.write "</tr>" 
	response.flush

	Do While Not oRs.EOF

		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr>"

		' Overpayment Type
		response.write "<td>"
		response.write oRs("overpayment")
		response.write "</td>"
		
		' Location
		response.write "<td>"
		response.write oRs("name")
		response.write "</td>"

		' Rental
		response.write "<td>"
		response.write oRs("rentalname")
		response.write "</td>"

		' Reservation #
		response.write "<td>"
		response.write oRs("reservationid")
		response.write "</td>"

		' Date 
		response.write "<td>"
		If oRs("overpayment") <> "Deposits and Charges" Then 
			response.write DateValue(CDate(oRs("reservationstarttime"))) 
		Else
			' get earliest start datetime and billing end datetime; reserved status before cancelled
			getEarliestDatesForReservation oRs("reservationid"), sStartDateTime, sEndBillingDateTime
			If sStartDateTime <> "" Then 
				response.write DateValue(CDate(sStartDateTime)) 
			End If 
		End If 
		response.write "</td>"

		' Time
		response.write "<td>"
		If oRs("overpayment") <> "Deposits and Charges" Then 
			response.write GetTimePortion( oRs("reservationstarttime") ) & " &ndash; " & GetTimePortion( oRs("billingendtime") )
		Else
			' use the datetimes retreived in the date display above
			If sStartDateTime <> "" And sEndBillingDateTime <> "" Then 
				response.write GetTimePortion( sStartDateTime ) & " &ndash; " & GetTimePortion( sEndBillingDateTime )
			End If 
		End If 
		response.write "</td>"

		' Rate
		response.write "<td align=""right"" class=""moneystyle"">"
		response.write oRs("rate")
		response.write "</td>"

		' Duration
		response.write "<td align=""right"" class=""moneystyle"">"
		If oRs("overpayment") = "Rental Fees" Then
			response.write CalculateDurationInHours( oRs("reservationstarttime"), oRs("billingendtime") )
		Else
			response.write oRs("duration")
		End If 
		response.write "</td>"

		' Fee Amount
		response.write "<td align=""right"" class=""moneystyle"">"
		response.write oRs("feeamount")
		response.write "</td>"

		' Paid Amount
		response.write "<td align=""right"" class=""moneystyle"">"
		response.write oRs("paidamount")
		response.write "</td>"

		' Refunded Amount
		response.write "<td align=""right"" class=""moneystyle"">"
		response.write oRs("refundamount")
		response.write "</td>"

		response.write "</tr>"
		response.flush

		oRs.MoveNext 
	Loop 

Else
	response.write vbcrlf & "<tr><td>Nothing was found matching your search criteria.</td></tr>"
	response.flush
End If 

response.write vbcrlf & "</table>"
response.write vbcrlf & "</body></html>"
response.flush

oRs.Close
Set oRs = Nothing 


Sub getEarliestDatesForReservation( ByVal iReservationId, ByRef sStartDateTime, ByRef sEndBillingDateTime )
	Dim oRs, sSql

	sStartDateTime = ""
	sEndBillingDateTime = ""

	sSql = "SELECT TOP 1 reservationstarttime, billingendtime FROM egov_rentalreservationdates where reservationid = " & iReservationId
	sSql = sSql & " ORDER BY statusid, reservationstarttime"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open  sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		sStartDateTime = oRs( "reservationstarttime" )
		sEndBillingDateTime = oRs( "billingendtime" )
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

%>
