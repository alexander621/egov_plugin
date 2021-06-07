<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationlistexport.asp
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
' 1.2	05/10/2010	Steve Loar - Added receipt number to search
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSql, oRs, oSchema, toDate, fromDate, sSearch, iRentalId, sRentalIdType, iActualId, iReservationTypeId
Dim sReservationTypeSelector, sRenterName, iReservedStatus, sStatusSearch, sReservationId, iSupervisorUserId
Dim sPaymentId, sFrom, sReportTitle, sLocation

sReportTitle = "<tr><td></td><td>Reservation List</td></tr>"

server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Rental_Reservations.xls"

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

sSearch = sSearch & " AND D.reservationstarttime BETWEEN '" & CDate(fromDate) & " 0:00 AM' AND '" & DateAdd("d",1,CDate(toDate)) &" 0:00 AM' "
sReportTitle = sReportTitle & "<tr><td>Reservation Start Time:</td><td>Between " & DateValue(CDate(fromDate)) & " AND " & DateValue(CDate(toDate)) & "</td></tr>"

If request("rentalid") <> "" Then
	iRentalId = request("rentalid")
	If iRentalId <> "0" Then 
		sRentalIdType = Left(iRentalId, 1)
		iActualId = Mid(iRentalId, 2)
		If sRentalIdType = "R" Then
			sSearch = sSearch & " AND D.rentalid = " & iActualId
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


If request("location") <> "" Then  
	sLocation = request("location")
	sSearch = sSearch & " AND ( RE.rentalname LIKE '%" & dbsafe(sLocation) & "%' OR L.name LIKE '%" & dbsafe(sLocation) & "%' ) "
	sReportTitle = sReportTitle & "<tr><td>Location Like:</td><td>" & sLocation & "</td></tr>"
End If 

If request("reservationtypeid") <> "" Then
	iReservationTypeId = CLng(request("reservationtypeid"))
	If iReservationTypeId <> CLng(0) Then 
		sSearch = sSearch & " AND R.reservationtypeid = " & iReservationTypeId
		sReservationTypeSelector = GetReservationTypeSelection( iReservationTypeId )  ' in rentalscommonfunctions.asp
		sReportTitle = sReportTitle & "<tr><td>Reservation Type:</td><td>" & GetReservationType( iReservationTypeId ) & "</td></tr>"
	Else
		sReservationTypeSelector = ""
		sReportTitle = sReportTitle & "<tr><td>Reservation Type:</td><td>All Reservation Types</td></tr>"
	End If 
Else
	iReservationTypeId = "0"
	sReservationTypeSelector = ""
	sReportTitle = sReportTitle & "<tr><td>Reservation Type:</td><td>All Reservation Types</td></tr>"
End If 

If request("rentername") <> "" Then
	sRenterName = request("rentername")
	sReportTitle = sReportTitle & "<tr><td>Renter Name Like: </td<td>" & request("rentername") & "</td></tr>"
Else
	sRenterName = ""
End If 

' Set up the choice of what statuses to show
If request("reservationstatus") = "" Then
	iReservedStatus = 1
Else
	iReservedStatus = request("reservationstatus")
End If 

Select Case iReservedStatus
	Case 1
		sStatusSearch = " AND DS.iscancelled = 0 AND RS.iscancelled = 0 AND R.isonhold = 0 "
		sReportTitle = sReportTitle & "<tr><td>Status:</td><td>Reserved Only</td></tr>"
	Case 2
		sStatusSearch = " "
		sReportTitle = sReportTitle & "<tr><td>Status:</td><td>Reserved, Cancelled, And On Hold</td></tr>"
	Case 3
		sStatusSearch = " AND DS.iscancelled = 1 "
		sReportTitle = sReportTitle & "<tr><td>Status:</td><td>Cancelled Only</td></tr>"
	Case 4
		sStatusSearch = " AND DS.iscancelled = 0 AND RS.iscancelled = 0 AND R.isonhold = 1 "
		sReportTitle = sReportTitle & "<tr><td>Status:</td><td>On Hold Only</td></tr>"
End Select 

If request("reservationid") <> "" Then
	sReservationId = request("reservationid")
	sSearch = sSearch & " AND R.reservationid = " & sReservationId
	sReportTitle = sReportTitle & "<tr><td>Reservation Id: </td><td>" & sReservationId & "</td></tr>"
Else
	sReservationId = ""
End If 


If request("supervisoruserid") <> "" Then
	iSupervisorUserId = CLng(request("supervisoruserid"))
	If iSupervisorUserId > CLng(0) Then
		sSearch = sSearch & " AND RE.supervisoruserid = " & iSupervisorUserId
	End If 
Else
	iSupervisorUserId = 0
End If 

If request("paymentid") <> "" Then
	sPaymentId = CLng(request("paymentid"))
	sSearch = sSearch & " AND J.reservationid = R.reservationid AND J.paymentid = " & sPaymentId
	sFrom = sFrom & ", egov_class_payment J "
	sReportTitle = sReportTitle & "<tr><td>Payment #: </td><td>" & sPaymentId & "</td></tr>"
End If 



sSql = "SELECT D.reservationid, D.reservationstarttime, D.billingendtime, RE.rentalname, L.name AS locationname, ISNULL(R.timeid,0) AS timeid, "
sSql = sSql & " (R.totalamount + R.totalrefunded - R.totalpaid + R.totalrefundfees) AS balancedue, R.totalamount, R.totalrefunded, R.isonhold, "
sSql = sSql & " RT.reservationtype, RT.reservationtypeselector, DS.reservationstatus, R.reserveddate, R.rentaluserid" & sSelect
sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservations R, egov_rentals RE, egov_class_location L, "
sSql = sSql & " egov_rentalreservationtypes RT, egov_rentalreservationstatuses DS, egov_rentalreservationstatuses RS" & sFrom
sSql = sSql & " WHERE D.reservationid = R.reservationid AND D.rentalid = RE.rentalid AND RE.locationid = L.locationid "
sSql = sSql & " AND R.reservationtypeid = RT.reservationtypeid AND D.statusid = DS.reservationstatusid AND RE.orgid = " & session("orgid")
sSql = sSql & " AND R.reservationstatusid = RS.reservationstatusid" & sStatusSearch & sSearch
sSql = sSql & " ORDER BY D.reservationstarttime, L.name, RE.rentalname"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

response.write vbcrlf & "<html>"

response.write vbcrlf & "<style>  "
response.write " .moneystyle "
response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
response.write vbcrlf & "</style>"

response.write vbcrlf & "<body><table border=""1"">"
'response.write "<tr><th></th><th>Reservations</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"
response.write sReportTitle

If Not oRs.EOF Then
	response.write "<tr><th>Reservation Id</th><th>Reservation Date</th><th>Time</th><th>Hours</th><th>Rental</th><th>Location</th><th>Type</th>"
	response.write "<th>Renter</th><th>Class Name</th><th>Status</th><th>Reserved</th>"
	'<th>Total<br />Fees</th><th>Amount<br />Due</th>
	response.write "</tr>"
	response.flush

	Do While Not oRs.EOF
		If oRs("reservationtypeselector") = "public" Then
			sRenter = GetCitizenName( oRs("rentaluserid") ) & "</td><td>"
			sRenterSearch = sRenter
		Else
			If oRs("reservationtypeselector") = "admin" Then
				sRenter = GetAdminName( oRs("rentaluserid") ) & "</td><td>"
				 sRenterSearch = sRenter
			Else 
				If oRs("reservationtypeselector") = "class" Then
					sRenter = GetActivityNo( oRs("timeid") )	' In rentalscommonfunctions.asp 
					sClassName = GetClassName( oRs("timeid") )	' In rentalscommonfunctions.asp 
					sRenterSearch = sRenter & " " & sClassName
					sRenter = sRenter & "</td><td>" & sClassName	
				Else
					' this leaves blocked
					sRenter = "" & "</td><td>"
					sRenterSearch = ""
				End If 
			End If 
		End If 

		' if the rentername is not empty then they are looking for a match of some sort
		If sRenterName <> "" Then 
			If oRs("reservationtypeselector") <> "block" Then
				' they are looking for someone. 0 is not found
				If InStr(LCase(sRenterSearch), LCase(sRenterName)) > 0 Then
					' the searched for name is in the name
					bOK = True 
				Else
					' the searched for name is not in the name
					bOK = False 
				End If 
			Else
				' we leave block out of name searches 
				bOk = False  
			End If 
		Else
			' No name search so the record is OK
			bOk = True 
		End If 

		If bOk Then 
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr>"

			' Reservation #
			response.write "<td>"
			response.write oRs("reservationid")
			response.write "</td>"

			' Date 
			response.write "<td>"
			response.write DateValue(CDate(oRs("reservationstarttime"))) 
			response.write "</td>"

			' Time
			response.write "<td>"
			response.write GetTimePortion( oRs("reservationstarttime") ) & " &ndash; " & GetTimePortion( oRs("billingendtime") )
			response.write "</td>"

			' Duration
			response.write "<td align=""right"" class=""moneystyle"">"
			response.write CalculateDurationInHours( oRs("reservationstarttime"), oRs("billingendtime") )
			response.write "</td>"

			' Location 
			response.write "<td>"
			response.write oRs("rentalname") 
			response.write "</td>"
			response.write "<td>"
			response.write oRs("locationname") 
			response.write "</td>"

			' Type
			response.write "<td>"
			response.write oRs("reservationtype")
			response.write "</td>"

			' Renter
			response.write "<td>"
			response.write sRenter
			response.write "</td>"

			' Status
			response.write "<td>"
			If oRs("isonhold") Then
				response.write "On Hold"
			Else 
				response.write oRs("reservationstatus")
			End If 
			response.write "</td>"

			' Reserved date
			response.write "<td>"
			response.write DateValue(oRs("reserveddate"))
			response.write "</td>"

			' Balance
'			If oRs("reservationtypeselector") = "public" Then 
'				response.write "<td>&nbsp;"
'				response.write FormatNumber(CDbl(oRs("totalamount")),2,,,0)
'				response.write "</td>"
'				response.write "<td>&mbsp;"
'				response.write FormatNumber(CDbl(oRs("balancedue")),2,,,0)
'				response.write "</td>"
'			Else
'				response.write "<td>"
'				response.write "&nbsp;</td>"
'				response.write "<td>"
'				response.write "&nbsp;"
'				response.write "</td>"
'			End If 

			response.write "</tr>"
			response.flush
		End If 

		oRs.MoveNext 
	Loop 

End If 

response.write "</table></body></html>"

oRs.Close
Set oRs = Nothing 




%>
