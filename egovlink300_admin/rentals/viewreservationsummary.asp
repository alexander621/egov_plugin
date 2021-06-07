<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: viewreservationsummary.asp
' AUTHOR: Steve Loar
' CREATED: 11/24/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays the summary information for a rental reservation
'
' MODIFICATION HISTORY
' 1.0	11/24/2009	Steve Loar - Initial Version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iReservationId, iReservationTypeId, sReservationType, sReservationStatus, sOrganization
Dim sPointOfContact, sNumberAttending, sReceiptNotes, sPrivateNotes, sReservedDate
Dim sReservationTypeSelector, sRenterName, sRenterPhone, sAdminName, sServingAlcohol
Dim bReservationIsCancelled, iRentalUserId, bIsReservation, sPurpose, iTimeid

sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "edit reservations", sLevel	' In common.asp

iReservationId = CLng(request("reservationid"))

GetGeneralReservationData iReservationId		' In rentalscommonfunctions.asp

%>
<html>
<head>
	<title>E-Gov Administration Console {Receipt}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="javascript">
	<!--

		window.onload = function()
		{
			//factory.printing.header = "Printed on &d"
			//factory.printing.footer       = "&bPrinted on &d - Page:&p/&P";
			//factory.printing.portrait     = true;
			//factory.printing.leftMargin   = 0.5;
			//factory.printing.topMargin    = 0.5;
			//factory.printing.rightMargin  = 0.5;
			//factory.printing.bottomMargin = 0.5;

			// enable control buttons
			//var templateSupported = factory.printing.IsTemplateSupported();
			//var controls = idControls.all.tags("input");
			//for ( i = 0; i < controls.length; i++ ) 
			//{
			//	controls[i].disabled = false;
			//	if (templateSupported && controls[i].className == "ie55" )
			//		controls[i].style.display = "inline";
			//}
		}

	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN: THIRD PARTY PRINT CONTROL-->
<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
<%
'	<input disabled type="button" value="Print the page" class="button" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
'	<input class="ie55" disabled type="button" value="Print Preview..." class="button" onclick="factory.printing.Preview()" />
%>
	&nbsp;&nbsp;
	<input type="button" class="button" value="<< Back To Reservation" onclick="location.href='reservationedit.asp?reservationid=<%=iReservationId%>';" />	
</div>

<%
'<object id="factory" viewastext  style="display:none"
'   classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
'   codebase="../includes/smsx.cab#Version=6,3,434,12">
'</object>
%>
<!--END: THIRD PARTY PRINT CONTROL-->

<!--BEGIN PAGE CONTENT-->
<div id="content">
 	<div id="centercontent">
<%
		ShowReceiptHeaderForOrg

		response.write vbcrlf & "<hr />"
		response.write "<span id=""receiptadmininfo"">"
		response.write " Reserved On: " & sReservedDate
		response.write "</span>"
		response.write " Reservation Id: " & iReservationId & "&nbsp;&nbsp;" 
		response.write " Status: " & sReservationStatus & "&nbsp;&nbsp;" 
		response.write " Type: " & sReservationType
		response.write vbcrlf & "<hr />" 

		If bIsReservation Then
			response.write vbcrlf & "<div id=""receipttopright"">"
			response.write vbcrlf & "<table id=""summaryeventinfo"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
			response.write vbcrlf & "<tr><td class=""summaryeventinfolabel"" valign=""top"">Organization:" & "</td><td valign=""top"">" & sOrganization & "</td></tr>"
			response.write vbcrlf & "<tr><td class=""summaryeventinfolabel"" valign=""top"">Point Of Contact:" & "</td><td valign=""top"">" & sPointOfContact & "</td></tr>"
			response.write vbcrlf & "<tr><td class=""summaryeventinfolabel"" valign=""top"">Number Attending:" & "</td><td valign=""top"">" & sNumberAttending & "</td></tr>"
			response.write vbcrlf & "<tr><td class=""summaryeventinfolabel"" valign=""top"">Purpose: " & "</td><td valign=""top"">" & sPurpose & "</td></tr>"
			response.write vbcrlf & "</table>"
			response.write vbcrlf & "</div>"
		 
			response.write "<div>"
			ShowRentalUserInfo iRentalUserId, sReservationTypeSelector, sRenterName, sRenterPhone 
			response.write "</div>"

			response.write vbcrlf & "<hr />" 
		Else
			If sReservationTypeSelector = "class" Then	
				response.write "<div><strong>"
				ShowClassActivityNoAndName iTimeId, False 
				response.write "</strong></div>"
				response.write vbcrlf & "<hr />" 
			End If 
		End If 

		response.write "<strong>Reservations</strong>"
		response.write vbcrlf & "<hr />" 

		ShowSummaryDatesAndFees iReservationId, sReservationTypeSelector, sServingAlcohol

%>

	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' ShowReceiptHeaderForOrg
'--------------------------------------------------------------------------------------------------
Sub ShowReceiptHeaderForOrg()

	response.write vbcrlf & "<p class=""receiptheader"">"
	If OrgHasDisplay( Session("orgid"), "rental receipt header" ) Then
		response.write GetOrgDisplay( Session("orgid"), "rental receipt header" ) 
		response.write vbcrlf & "<br /><br />"
	End If 
	response.write "Reservation Summary"
	response.write vbcrlf & "</p>"

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowRentalUserInfo iRentalUserId, sReservationTypeSelector 
'--------------------------------------------------------------------------------------------------
Sub ShowRentalUserInfo( ByVal iRentalUserId, ByVal sReservationTypeSelector, ByVal sRenterName, ByVal sRenterPhone )
	Dim sUserDisplay

	response.write vbcrlf & "<span class=""receipttitles"">"
	response.write "Renter Information</span><br />"
	'response.write iRentalUserId & " " & sReservationTypeSelector & "<br />"
	If sReservationTypeSelector = "public" Then
		' Citizen Renter
		sUserDisplay = GetRentalCitizenInfo( iRentalUserId )
		response.write vbcrlf & sUserDisplay
	Else
		' Admin Renter
		response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""receiptuserinfo"">"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td nowrap=""nowrap""><strong>" & sRenterName & "</strong></td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td>" & sRenterPhone & "</td></tr>"
		response.write vbcrlf & "</table>"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' string sInfo = GetRentalCitizenInfo( iRentalUserId )
'--------------------------------------------------------------------------------------------------
Function GetRentalCitizenInfo( ByVal iRentalUserId )
	Dim oCmd, oRs, sDisplayText

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserInfoList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iuserid", 3, 1, 4, iRentalUserId)
	    Set oRs = .Execute
	End With

	sDisplayText = sDisplayText & vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""receiptuserinfo"">"
	sDisplayText = sDisplayText & vbcrlf & "<tr><td align=""right"" valign=""top"">&nbsp;</td><td nowrap=""nowrap""><strong>" & oRs("userfname") & " " & oRs("userlname") & "</strong><br />"
	sDisplayText = sDisplayText & "<strong>" & oRs("useraddress") 
	If oRs("userunit") <> "" Then 
		sDisplayText = sDisplayText & "&nbsp;&nbsp;" & oRs("userunit") 
	End If
	If oRs("useraddress2") <> "" Then 
		sDisplayText = sDisplayText & "<br />" & oRs("useraddress2") 
	End If 
	sDisplayText = sDisplayText & "<br />" & oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") & "</strong></td></tr>"
	sDisplayText = sDisplayText & vbcrlf & "<tr><td colspan=""2"">&nbsp;</td></tr>"
	sDisplayText = sDisplayText & vbcrlf & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & GetFamilyEmail( iRentalUserId ) & "</td></tr>"
	sDisplayText = sDisplayText & vbcrlf & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhoneNumber(oRs("userhomephone")) & "</td></tr>"
	sDisplayText = sDisplayText & vbcrlf & "</table>"

	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing

	GetRentalCitizenInfo = sDisplayText
End Function 


'--------------------------------------------------------------------------------------------------
' ShowSummaryDatesAndFees iReservationId, sReservationTypeSelector, sServingAlcohol, bReservationIsCancelled
'--------------------------------------------------------------------------------------------------
Sub ShowSummaryDatesAndFees( ByVal iReservationId, ByVal sReservationTypeSelector, ByVal sServingAlcohol )
	Dim sSql, oRs, iRowCount, sStartHour, sStartMinute, sStartAmPm, sEndHour, sEndMinute, sEndAmPm
	Dim sAmPm, iDateTotal, iReservationFeeCount, iRentalRateCount, iReservationItemCount, dTotalRefundFees
	Dim iReservationDateCount, dTotalRefunded, dTotalPaid, bIsCancelled, dTotalRefundCombined, iTotalDue
	Dim sArrivalHour, sArrivalMinute, sArrivalAmPm, sDepartureHour, sDepartureMinute, sDepartureAmPm

	iRowCount = 0
	iReservationFeeCount = clng(0)
	iRentalRateCount = clng(0)
	iReservationItemCount = clng(0)
	iReservationDateCount = clng(0)
	iTotalDue = CDbl(0.00)

	' Get the reserved dates
	sSql = "SELECT D.reservationdateid, D.reservationid, D.reservationstarttime, D.billingendtime, D.actualstarttime, "
	sSql = sSql & " D.actualendtime, D.reserveddate, D.adminuserid, D.rentalid, S.reservationstatus, S.iscancelled "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservationstatuses S "
	sSql = sSql & " WHERE D.statusid = S.reservationstatusid AND D.orgid = " & session("OrgId")
	sSql = sSql & " AND D.reservationid = " & iReservationId
	sSql = sSql & " ORDER BY reservationstarttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<table id=""datesandfees"" cellpadding=""2"" cellspacing=""0"" border=""0"">"

	' Show the daily fees and items
	Do While Not oRs.EOF
		' Display the date info for all
		iRowCount = iRowCount + 1
		iReservationDateCount = iReservationDateCount + 1
		iDateTotal = CDbl(0.00)
		
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & ">"

		' Date
		response.write "<td><strong>"
		response.write DateValue(oRs("reservationstarttime")) & " &nbsp;(" & WeekDayName(Weekday(DateValue(oRs("reservationstarttime")))) & ")"
		response.write "</strong></td>"

		' From and to times
		response.write "<td><strong>" & FormatTimeString( oRs("reservationstarttime") )
		response.write " &mdash; " & FormatTimeString( oRs("billingendtime") )
		sDuration = CalculateDurationInHours( oRs("reservationstarttime"), oRs("billingendtime") )
		'response.write "</strong> (" &  & " hours)"
		response.write "</td>"
		
		' Location
		response.write "<td>"
		ShowRentalNameAndLocation oRs("rentalid") 
		response.write "</td>"

		response.write "<td>&nbsp;</td>"

		response.write "</tr>"

		response.write vbcrlf & "<tr" & sClass & ">"

		' Status and Cancel button
		response.write "<td valign=""top""><strong>"
		response.write oRs("reservationstatus") & "</strong>"
		If Not oRs("iscancelled") Then
			bIsCancelled = False 
		Else
			bIsCancelled = True 
		End If 
		response.write "</td>"

		' Arrival and departure times 
		
		'response.write "<input type=""hidden"" name=""reservationdateid" & iReservationDateCount & """ value=""" & oRs("reservationdateid") & """ />"
		'response.write "<input type=""hidden"" name=""reservationarrivaldate" & iReservationDateCount & """ value=""" & DateValue(oRs("reservationstarttime")) & """ />"
		'response.write "<input type=""hidden"" name=""reservationdeparturedate" & iReservationDateCount & """ value=""" & DateValue(oRs("billingendtime")) & """ />"
		If sReservationTypeSelector <> "block" And sReservationTypeSelector <> "class" Then
			' These are public and internal types
			response.write "<td valign=""top"">"
			response.write "Arrival: " & FormatTimeString( oRs("actualstarttime") )
			response.write "&nbsp;"

			response.write "Departure:" & FormatTimeString( oRs("actualendtime") )
			response.write "</td>"
		Else
			response.write "<td>&nbsp;</td>"
		End If 

		response.write "<td colspan=""2"">&nbsp;</td>"
		response.write "</tr>"

		If sReservationTypeSelector <> "block" And sReservationTypeSelector <> "class" Then
			' Display the hourly rates for reservations
			response.write vbcrlf & "<tr" & sClass & "><td colspan=""3"" align=""right"">Charge for " & sDuration & " hrs</td>"
			response.write "<td colspan=""2"" align=""right"">" & GetRentalRatesForDate( oRs("reservationdateid"), sClass, iTotalDue, iDateTotal ) & "</td>"
			response.write "</tr>"
			
		End If 

		If sReservationTypeSelector <> "block" Then 
			' Display items for all but blocked
			ShowItemsForDate oRs("reservationdateid"), sReservationTypeSelector, sClass, iTotalDue, iDateTotal
		End If 

		If sReservationTypeSelector <> "block" And sReservationTypeSelector <> "class" Then
			' Sub Total Row
			response.write vbcrlf & "<tr" & sClass & "><td class=""subtotalscell"" colspan=""3"" align=""right""><strong>Subtotal</strong></td>"
			response.write "<td class=""subtotalscell"" align=""right"">"
			response.write FormatNumber(iDateTotal,2,,,0) 
			response.write "</td></tr>"
		End If 

		oRs.MoveNext
	Loop

	' Reservation Fees 
	If sReservationTypeSelector = "public" Then
		' Display Reservation Level charges like Deposit and Alcohol Fee
		ShowReservationFees iReservationId, iRowCount, iTotalDue, sServingAlcohol
'	Else
'		response.write vbcrlf & "<tr" & sClass & "><td colspan=""4"">&nbsp;</td></tr>"
	End If 

	If sReservationTypeSelector <> "block" And sReservationTypeSelector <> "class" Then
		' Total Charges Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""2"" align=""right"">&nbsp;"
		response.write "</td>"
		response.write "<td class=""totalscell"" align=""right""><strong>Total Charges</strong></td>"
		response.write "<td class=""totalscell"" align=""right"">"
		response.write FormatNumber(iTotalDue,2,,,0) 
		response.write "</td></tr>"

		' Total Paid Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""4"" align=""center""><strong>Payments</strong></td></tr>"
		' Want to show payments with link to receipt page here
		ShowReservationPayments iReservationId, sClass
		response.write vbcrlf & "<tr" & sClass & ">"
		response.write "<td colspan=""3"" align=""right"" class=""totalscell""><strong>Total Paid</strong></td>"
		response.write "<td align=""right"" class=""totalscell"">"
		'dTotalPaid = GetReservationTotalAmount( iReservationId, "totalpaid" ) ' In rentalscommonfunctions.asp
		dTotalPaid = GetTotalPaidForReservation( iReservationId )  ' In rentalscommonfunctions.asp
		response.write FormatNumber(dTotalPaid,2,,,0) 
		response.write "</td></tr>"

		' Refund Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 

		If sReservationTypeSelector = "public" Then
			response.write vbcrlf & "<tr" & sClass & "><td class=""totalscell"" colspan=""4"" align=""center""><strong>Refunds</strong></td></tr>"
			' Want to show refunds with link to receipt page here
			ShowReservationRefunds iReservationId, sClass
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td class=""totalscell"" colspan=""2"">&nbsp;</td>"
			response.write "<td class=""totalscell"" align=""right""><strong>Total Refunds</strong></td>"
			response.write "<td class=""totalscell"" align=""right"">"
			dTotalRefunded = GetReservationTotalAmount( iReservationId, "totalrefunded" ) ' In rentalscommonfunctions.asp - is just the pull of the total field
			response.write FormatNumber(dTotalRefunded,2,,,0) 
			response.write "</td></tr>"
		End If 

		' Balance Due Row
		iRowCount = iRowCount + 1
		If iRowCount Mod 2 = 0 Then
			sClass = " class=""altrow"" "
		Else
			sClass = ""
		End If 
		response.write vbcrlf & "<tr" & sClass & ">"
		response.write "<td class=""totalscell"" colspan=""3"" align=""right""><strong>Balance Due</strong></td>"
		response.write "<td class=""totalscell"" align=""right"">"
		dBalanceDue = (iTotalDue + dTotalRefunded) - dTotalPaid
		response.write FormatNumber(dBalanceDue,2,,,0) 
		response.write "</td></tr>"
	End If 

	response.write vbcrlf & "</table>"

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' string sTotalRate = GetRentalRatesForDate iReservationDateId, sClass, iTotalDue, iDateTotal
'--------------------------------------------------------------------------------------------------
Function GetRentalRatesForDate( ByVal iReservationDateId, ByVal sClass, ByRef iTotalDue, ByRef iDateTotal )
	Dim sSql, oRs, iRateTotal

	sSql = "SELECT F.reservationdatefeeid, F.amount, F.feeamount, F.duration, P.pricetypename, F.starthour, "
	sSql = sSql & " dbo.AddLeadingZeros(F.startminute,2) AS startminute, F.startampm, P.isweekendsurcharge, R.ratetype "
	sSql = sSql & " FROM egov_rentalreservationdatefees F, egov_price_types P, egov_rentalratetypes R "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.ratetypeid = R.ratetypeid AND F.reservationdateid = " & iReservationDateId
	sSql = sSql & " ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRateTotal = iRateTotal +  CDbl(oRs("feeamount"))
		iTotalDue = iTotalDue + CDbl(oRs("feeamount"))
		iDateTotal = iDateTotal + CDbl(oRs("feeamount"))
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	GetRentalRatesForDate = FormatNumber(iRateTotal,2,,,0)
End Function 


'--------------------------------------------------------------------------------------------------
' ShowItemsForDate iReservationDateId, sReservationTypeSelector, sClass, iTotalDue, iDateTotal
'--------------------------------------------------------------------------------------------------
Sub ShowItemsForDate( ByVal iReservationDateId, ByVal sReservationTypeSelector, ByVal sClass, ByRef iTotalDue, ByRef iDateTotal )
	Dim oRs, sSql

	sSql = "SELECT reservationdateitemid, rentalitem, ISNULL(quantity,0) AS quantity, "
	sSql = sSql & " ISNULL(feeamount,0.00) AS feeamount "
	sSql = sSql & " FROM egov_rentalreservationdateitems "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND quantity IS NOT NULL "
	sSql = sSql & " ORDER BY rentalitem"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If clng(oRs("quantity")) > clng(0) Then 
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td colspan=""3"" align=""right"">"
			response.write FormatNumber(oRs("quantity"),0,,,0) & "&nbsp;" & oRs("rentalitem")
			response.write "</td>"
			response.write "<td align=""right"">"
			If sReservationTypeSelector <> "block" And sReservationTypeSelector <> "class" Then
				response.write FormatNumber(oRs("feeamount"),2,,,0)
				iTotalDue = iTotalDue + CDbl(oRs("feeamount"))
				iDateTotal = iDateTotal + CDbl(oRs("feeamount"))
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"
			response.write "</tr>"
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowReservationFees iReservationId, iRowCount, iTotalDue, sServingAlcohol
'--------------------------------------------------------------------------------------------------
Sub ShowReservationFees( ByVal iReservationId, ByRef iRowCount, ByRef iTotalDue, ByVal sServingAlcohol)
	Dim oRs, sSql, sCellClass, iReservationFeeCount

	iReservationFeeCount = clng(0)

	sSql = "SELECT F.reservationfeeid, P.pricetypename, F.amount, F.feeamount, ISNULL(F.prompt,'') AS prompt, P.isalcoholsurcharge "
	sSql = sSql & " FROM egov_rentalreservationfees F, egov_price_types P "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.reservationid = " & iReservationId
	sSql = sSql & " ORDER BY P.displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If CDbl(oRs("feeamount")) > CDbl(0.00) Then 
			iReservationFeeCount = iReservationFeeCount + clng(1)
			If iReservationFeeCount = clng(1) Then
				iRowCount = iRowCount + 1
				If iRowCount Mod 2 = 0 Then
					sClass = " class=""altrow"" "
				Else
					sClass = ""
				End If 
				sCellClass = " class=""totalscell"""
			Else
				sCellClass = ""
			End If 
			response.write vbcrlf & "<tr" & sClass & ">"
			response.write "<td colspan=""3"" align=""right""" & sCellClass & ">"
			If oRs("isalcoholsurcharge") Then
				' There will only be one per reservation so no need to put the count on this
				response.write "<input type=""checkbox"" name=""servingalcohol""" & sServingAlcohol
				'response.write " disabled=""disabled"" "
				response.write " /> " & oRs("prompt") & " &nbsp; &mdash; &nbsp; "
			End If 
			response.write oRs("pricetypename")
			response.write "</td>"
			response.write "<td align=""right""" & sCellClass & ">"
			response.write FormatNumber(oRs("feeamount"),2,,,0)
			iTotalDue = iTotalDue + CDbl(oRs("feeamount"))
			response.write "</td>"
			response.write "</tr>"
		End If 
		oRs.MoveNext 
	Loop

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowReservationPayments iReservationId, sClass
'--------------------------------------------------------------------------------------------------
'Sub ShowReservationPayments( ByVal iReservationId, ByVal sClass )
'	Dim sSql, oRs

'	sSql = "SELECT A.paymentid, J.paymentdate, SUM(A.amount) AS paidamount "
'	sSql = sSql & "FROM egov_accounts_ledger A, egov_class_payment J "
'	sSql = sSql & "WHERE A.paymentid = J.paymentid AND A.ispaymentaccount = 1 AND A.entrytype = 'debit' "
'	sSql = sSql & "AND A.reservationid = " & iReservationId
'	sSql = sSql & "GROUP BY A.paymentid, J.paymentdate ORDER BY J.paymentdate"

'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSql, Application("DSN"), 0, 1

'	If Not oRs.EOF Then 
'		response.write vbcrlf & "<tr" & sClass & "><td class=""subheadercell"">Receipt #</td><td class=""subheadercell"">Date</td><td class=""subheadercell"">&nbsp;</td><td align=""right"" class=""subheadercell"">Amount</td></tr>"
'		Do While Not oRs.EOF
'			response.write vbcrlf & "<tr" & sClass & ">"
'			response.write "<td>" & oRs("paymentid") & "</td>"
'			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>"
'			response.write "<td>&nbsp;</td>"
'			response.write "<td align=""right"">"
'			response.write FormatNumber(oRs("paidamount"),2,,,0) 
'			response.write "</td></tr>"
'			oRs.MoveNext 
'		Loop
'	End If 
	
'	oRs.Close
'	Set oRs = Nothing 
'End Sub 


'--------------------------------------------------------------------------------------------------
' ShowReservationRefunds iReservationId, sClass
'--------------------------------------------------------------------------------------------------
'Sub ShowReservationRefunds( ByVal iReservationId, ByVal sClass )
'	Dim sSql, oRs
	
'	sSql = "SELECT A.paymentid, J.paymentdate, SUM(A.amount) as refundamount "
'	sSql = sSql & " FROM egov_accounts_ledger A, egov_class_payment J "
'	sSql = sSql & " WHERE A.paymentid = J.paymentid AND A.ispaymentaccount = 0 AND A.entrytype = 'debit' "
'	sSql = sSql & " AND A.reservationid = " & iReservationId
'	sSql = sSql & " GROUP BY A.paymentid, J.paymentdate ORDER BY J.paymentdate"

'	Set oRs = Server.CreateObject("ADODB.Recordset")
'	oRs.Open sSql, Application("DSN"), 0, 1

'	If Not oRs.EOF Then 
'		response.write vbcrlf & "<tr" & sClass & "><td class=""subheadercell"">Receipt #</td><td class=""subheadercell"" colspan=""2"">Date</td><td align=""right"" class=""subheadercell"">Amount</td></tr>"
'		Do While Not oRs.EOF
'			response.write vbcrlf & "<tr" & sClass & ">"
'			response.write "<td>&nbsp;" & oRs("paymentid") & "</td>"
'			response.write "<td>" & DateValue(oRs("paymentdate")) & "</td>"
'			response.write "<td>&nbsp;</td>"
'			response.write "<td align=""right"">"
'			response.write FormatNumber(oRs("refundamount"),2,,,0) 
'			response.write "</td></tr>"
'			oRs.MoveNext
'		Loop
'	End If 
	
'	oRs.Close
'	Set oRs = Nothing 

'End Sub 
%>
