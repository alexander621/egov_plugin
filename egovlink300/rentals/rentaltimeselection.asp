<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaltimeselection.asp
' AUTHOR: Steve Loar
' CREATED: 01/28/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Selection of specific times for rental reservations
'
' MODIFICATION HISTORY
' 1.0   01/28/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Force the page to be re-loaded on back button
response.Expires = 60
response.Expiresabsolute = Now() - 1
response.AddHeader "pragma","no-store"
response.AddHeader "cache-control","private"
response.CacheControl = "no-store" 'HTTP prevent back button after purchase problems

Dim iReservationTempId, iRentalId, bHasData, bHasHours, sSelectedDate, bOffSeasonFlag, sTitle
Dim bIsAllDayOnly, sStartTimeLabel, sEndTimeLabel, sIsAllDay, sMessage, bInitial, sResidentType
Dim iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm, iCitizenUserId
Dim iArrivalHour, iArrivalMinute, sArrivalAmPm, iDepartureHour, iDepartureMinute, sDepartureAmPm
Dim iMaxCharges, iIncludePriceTypeId, sLatestStart, bOkToDisplay, dNonResidentStartDate, sLessThanDate
Dim bFutureProblem, bNotOfferedToNonresidents

If request("rti") = "" Then
	response.redirect "rentalcategories.asp"
Else 
	If Not IsNumeric(request("rti")) Then
		response.redirect "rentalcategories.asp"
	Else 
		iReservationTempId = CLng(request("rti"))
	End If 
End If 

If request("pk") <> "" Then
	bInitial = True 
Else 
	bInitial = False 
End If 

iStartHour = "1"
iStartMinute = "00"
sStartAmPm = "PM"
iEndHour = "2"
iEndMinute = "00"
sEndAmPm = "PM"
iArrivalHour = "1"
iArrivalMinute = "00"
sArrivalAmPm = "PM"
iDepartureHour = "2"
iDepartureMinute = "00"
sDepartureAmPm = "PM"
iCitizenUserId = 0
iMaxCharges = 0
iIncludePriceTypeId = 0
sLatestStart = ""
bFutureProblem = False 
bNotOfferedToNonresidents = False

' still need to confirm that the data is there, and if not take them away from this page.
bHasData = SetPageVariables( iReservationTempId, iOrgId )

If bHasData = False  Then 
	' Take them somewhere safe, as their data is gone.
	response.redirect "rentalcategories.asp"
End If 

sResidentType = GetUserResidentType( iCitizenUserId )
'If they are not one of these (R, N), we have to figure which they are
If sResidentType <> "R" And sResidentType <> "N" and sResidentType <> "Z" Then 
	'This leaves E and B - See if they are a resident, also
	sResidentType = GetResidentTypeByAddress( iCitizenUserId, iOrgId )
End If 

If sResidentType = "R" Then
	' Check if the wanted date is too far out for any limits set on this rental
	If DateIsPastAllowedRange( iRentalId, sSelectedDate, "R", sLessThanDate ) Then
		bOkToDisplay = False
		bFutureProblem = True 
	Else 
		bOkToDisplay = True 
	End If 
Else
	' Handle Non-residents
	' First check that non-resident, or everyone pricing is offered for this rental on the selected day of the week
	If NonresidentsCanPurchase( iRentalId, Weekday(CDate(sSelectedDate)) ) Then 
		bNotOfferedToNonresidents = False
		' next check for any wanted date that is too far out for any limits set on this rental
		If DateIsPastAllowedRange( iRentalId, sSelectedDate, "N", sLessThanDate ) Then
			bOkToDisplay = False
			bFutureProblem = True 
		Else
			' else if no limits on how far out is breached then check the inseason only and non-res wait period
			If RentalHasInSeasonOnly( iRentalId ) And RentalHasNonResidentWait( iRentalId ) Then
				dNonResidentStartDate = GetNonResidentStartDate( iRentalId )
				If CDate(DateValue(Date())) < dNonResidentStartDate Then
					bOkToDisplay = False 
				Else
					bOkToDisplay = True
				End If 
			Else
				bOkToDisplay = True
			End If 
		End If 
	Else
		' Non-residents cannot make reservations at this facility
		bOkToDisplay = False
		bNotOfferedToNonresidents = True 
	End If 
End If 

'sLoadMsg = "displayScreenMsg('sLessThanDate: " & sLessThanDate & "');"

sMessage = request("msg")
If sMessage = "st" Then
	sLoadMsg = "displayScreenMsg('Check Failed: The time period you selected is less than the allowed minimum time for this location.');"
End If 
If sMessage = "conflict" Then
	sLoadMsg = "displayScreenMsg('Check Failed: The time period you selected is not available for reservations.');"
End If 
If sMessage = "closed" Then
	sLoadMsg = "displayScreenMsg('Check Failed: The time period you selected is beyond the open hours of this location.');"
End If 
If sMessage = "toolate" Then
	sLoadMsg = "displayScreenMsg('Check Failed: The time period you selected starts too late in the day for this location.');"
End If 
If sMessage = "sm" Then
	sLoadMsg = "displayScreenMsg('Check Failed: The times you selected are the same time. Please select different times and try again.');"
End If 


If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If


%>

<html lang="en">
<head>
  	<meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />
	<meta charset="UTF-8">

	<title><%=sTitle%></title>

	<link rel="stylesheet" href="../css/styles.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="rentalstyles.css" />
	<link rel="stylesheet" href="../css/style_<%=iorgid%>.css" />

	<script src="../scripts/jquery-1.7.2.min.js"></script>
	<style>
		.button,p#checktimes>input.button
		{
			padding:3px 8px !important;
		}
	</style>

	<script>
	<!--
	
		function goBack()
		{
			document.frmBack.submit();
		}

		function checkTime()
		{
			document.frmRentalTime.submit();
		}

		function displayScreenMsg( screenMsg ) 
		{
			if( screenMsg != "" ) 
			{
				$( "#screenMsg" ).html( screenMsg );
				window.setTimeout( "clearScreenMsg()", (10 * 1000) );
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html( "" );
		}

<%		If sLoadMsg <> "" Then	%>		
			
			$( document ).ready(
				function () {
					<%=sLoadMsg%>
				});

<%		End If %>


	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->
<%if iorgid = "228" and sResidentType = "Z" then response.redirect  "../manage_account.asp"%>

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>

<!--BEGIN: Page Top Display-->
<% 
	If OrgHasDisplay( iorgid, "rentalscategorypagetop" ) Then
		response.write vbcrlf & "<div id=""rentalscategorypagetop"">" & GetOrgDisplay( iOrgId, "rentalscategorypagetop" ) & "</div>"
	End If 
%>
<!--END: Page Top Display-->

<span id="screenMsg">&nbsp;</span>

<p>
	<input type="button" class="button" value="<< Back" onclick="goBack();" />
</p>

<%	ShowRentalsDetails iRentalId		%>

<% If bOkToDisplay Then %>

<p>Selected Date: <span id="pickeddate"><% response.write WeekDayName(Weekday(CDate(sSelectedDate))) & ", " & sSelectedDate%></span></p>

<p><strong>Available:</strong><br />
<% 
	
	bOffSeasonFlag = GetOffSeasonFlag( iRentalId, CDate(sSelectedDate) )
	bHasHours = DisplayAvailability( iRentalId, CDate(sSelectedDate), bOffSeasonFlag )
	response.write "</p>"

	If bHasHours Then 
		bIsAllDayOnly = ShowMinimumReservationTime( iRentalid, bOffSeasonFlag, Weekday(CDate(sSelectedDate)) )
	Else
		bIsAllDayOnly = False 
	End If 

	If bIsAllDayOnly Then
		sStartTimeLabel = "Arrival"
		sEndTimeLabel = "Departure"
		sIsAllDay = "1"
		If bInitial Then
			' Set times to opening and closing times
			SetTimeToOpeningAndClosing iRentalid, bOffSeasonFlag, Weekday(CDate(sSelectedDate)), CDate(sSelectedDate), iFirstHour, iFirstMinute, iFirstAmPm, iLastHour, iLastMinute, iLastAmPm
		Else 
			iFirstHour = iArrivalHour
			iFirstMinute = iArrivalMinute
			iFirstAmPm = sArrivalAmPm
			iLastHour = iDepartureHour
			iLastMinute = iDepartureMinute
			iLastAmPm = sDepartureAmPm
		End If 
	Else
		' Show Latest Starting time
		sLatestStart = GetLatestReservationHour( iRentalid, bOffSeasonFlag, Weekday(CDate(sSelectedDate)) )
		If sLatestStart <> "" Then 
			response.write "<p>The latest starting time for a reservation is " & sLatestStart & "</p>"
		End If 

		sStartTimeLabel = "Starting"
		sEndTimeLabel = "Ending"
		sIsAllDay = "0"
		If bInitial Then
			' Set times to first available and the end time to the minium duration out
			SetTimeToFirstAvailable iRentalid, bOffSeasonFlag, Weekday(CDate(sSelectedDate)), CDate(sSelectedDate), iFirstHour, iFirstMinute, iFirstAmPm, iLastHour, iLastMinute, iLastAmPm
		Else
			iFirstHour = iStartHour
			iFirstMinute = iStartMinute
			iFirstAmPm = sStartAmPm
			iLastHour = iEndHour
			iLastMinute = iEndMinute
			iLastAmPm = sEndAmPm
		End If 
	End If 

'	If bInitial Then
		' Set times to first available and the end time to the minium duration out
'	End If 

%>
</p>

<form method="post" name="frmRentalTime" action="rentalcontrol.asp">
	<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
	<input type="hidden" id="isallday" name="isallday" value="<%=sIsAllDay%>" />
	<input type="hidden" id="src" name="src" value="ts" />

<div id="ratesandchargesgroup">
<p><strong>Rates and Charges:</strong><br />
<% 
	If RentalHasNoCosts( iRentalId ) Then
		response.write "<strong>There are no charges for this reservation.</strong>"
	Else 
		response.write vbcrlf & "<table id=""ratesandcharges"" cellpadding=""2"" cellspacing=""0"" border=""0"">"

		ShowRates iRentalid, bOffSeasonFlag, Weekday(CDate(sSelectedDate)), sResidentType

		iMaxCharges = ShowRentalCharges( iRentalid, iIncludePriceTypeId )

		response.write vbcrlf & "</table>"
		response.write "<input type=""hidden"" id=""maxrentalcharges"" name=""maxrentalcharges"" value=""" & iMaxCharges & """ />"
	End If 
%>
</p>
</div>
	
	<p><strong>Select the <%=LCase(sStartTimeLabel)%> and <%=LCase(sEndTimeLabel)%> times for this reservation:</strong><br /><br />
			<span class="respCol">
<%			response.write "<strong>" & sStartTimeLabel & " Time</strong> &nbsp; "
			ShowHourPicks LCase(sStartTimeLabel) & "hour",  iFirstHour, "" 
			response.write ":"
			ShowMinutePicks LCase(sStartTimeLabel) & "minute", iFirstMinute, "" 
			response.write "&nbsp;"
			ShowAmPmPicks LCase(sStartTimeLabel) & "ampm", iFirstAmPm, "" 
			response.write " &nbsp;&nbsp;&nbsp; "
			response.write "</span>"
			response.write "<span class=""respCol"">"
			response.write "<strong>" & sEndTimeLabel & " Time</strong> &nbsp; "
			ShowHourPicks LCase(sEndTimeLabel) & "hour",  iLastHour, "" 
			response.write ":"
			ShowMinutePicks LCase(sEndTimeLabel) & "minute", iLastMinute, "" 
			response.write "&nbsp;"
			ShowAmPmPicks LCase(sEndTimeLabel) & "ampm", iLastAmPm, "" 
			response.write "</span>"

			ShowOrgRoundUpTime iOrgId
%>
	</p>

	<p id="checktimes">
		<input type="button" class="button" value="Check Times and Continue Reservation" onclick="checkTime();" />
	</p>

</form>

<%

Else
	' It is not ok to display
	If sResidentType = "R" Then
		' They tried for a date too far into the future
		response.write "<p id=""notreservablemsg"">This location limits residents to reservations that are before " & DateValue(sLessThanDate) & ". <br />The date you selected was " & DateValue(sSelectedDate) & ".</p>"
	Else 
		' Non resident with a date too far in the future or a wait period for making reservations
		If bFutureProblem Then
			response.write "<p id=""notreservablemsg"">This location limits non-residents to reservations that are before " & DateValue(sLessThanDate) & ". <br />The date you selected was " & DateValue(sSelectedDate) & ".</p>"
		Else 
			If bNotOfferedToNonresidents Then
				' Non-resident pricing is not offered
				response.write "<p id=""notreservablemsg"">We're sorry, but reservations at this facility are not available to Non-Residents.</p>"
			Else 
				' Non-residents have to wait until after a specific calendar date to start making reservations
				response.write "<p id=""notreservablemsg"">Non-Residents cannot make reservations until " & dNonResidentStartDate & ".</p>"
			End If 
		End If 
	End If 
End If 

%>


<form name="frmBack" method="post" action="rentalavailability.asp">
	<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
</form>


<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' boolean SetPageVariables( iReservationTempId, iOrgId )
'--------------------------------------------------------------------------------------------------
Function SetPageVariables( ByVal iReservationTempId, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT rentalid, selecteddate, ISNULL(starthour,1) AS starthour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(startminute,0),2) AS startminute, "
	sSql = sSql & " ISNULL(startampm,'PM') AS startampm, ISNULL(endhour,2) AS endhour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(endminute,0),2) AS endminute, "
	sSql = sSql & " ISNULL(endampm,'PM') AS endampm, ISNULL(arrivalhour,1) AS arrivalhour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(arrivalminute,0),2) AS arrivalminute, "
	sSql = sSql & " ISNULL(arrivalampm,'PM') AS arrivalampm, ISNULL(departurehour,2) AS departurehour, "
	sSql = sSql & " dbo.AddLeadingZeros(ISNULL(departureminute,0),2) AS departureminute, "
	sSql = sSql & " ISNULL(departureampm,'PM') AS departureampm, ISNULL(citizenuserid,0) AS citizenuserid, "
	sSql = sSql & " ISNULL(includepricetypeid,0) AS includepricetypeid "
	sSql = sSql & " FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	sSql = sSql & " AND orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iRentalId = CLng(oRs("rentalid"))
		sSelectedDate = oRs("selecteddate")
		iStartHour = oRs("starthour")
		iStartMinute = oRs("startminute")
		sStartAmPm = oRs("startampm")
		iEndHour = oRs("endhour")
		iEndMinute = oRs("endminute")
		sEndAmPm = oRs("endampm")
		iArrivalHour = oRs("arrivalhour")
		iArrivalMinute = oRs("arrivalminute")
		sArrivalAmPm = oRs("arrivalampm")
		iDepartureHour = oRs("departurehour")
		iDepartureMinute = oRs("departureminute")
		sDepartureAmPm = oRs("departureampm")
		iCitizenUserId = oRs("citizenuserid")
		iIncludePriceTypeId = oRs("includepricetypeid")
		SetPageVariables = True 
	Else
		SetPageVariables = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowRentalsDetails iRentalId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalsDetails( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, ISNULL(R.width,'') AS width, ISNULL(R.length,'') AS length, "
	sSql = sSql & "ISNULL(R.capacity,'') AS capacity, R.publiccanreserve, usehtmlonlongdesc, "
	sSql = sSql & "ISNULL(R.description,'') AS description, ISNULL(R.iconimageurl,'') AS iconimageurl "
	sSql = sSql & "FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.publiccanview = 1 AND R.locationid = L.locationId "
	sSql = sSql & "AND R.rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		'response.write vbCrLf & "<div class=""rentalfacilityname"">" & oRs("rentalname") & "</div>" 
		response.write "<table class=""availablerentals"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write "<tr><td class=""spacerrow"" colspan=""2"">&nbsp;</td></tr>"
		response.write "<tr>"
		response.write "<td valign=""top"" class=""iconcell"">" 
		If oRs("iconimageurl") <> "" Then 
				response.write "<img src=""" & replace(oRs("iconimageurl"),"http://www.egovlink.com","") & """ alt=""" & oRs("rentalname") & """ title=""" & oRs("rentalname") & """ class=""availabilityimg"" />"
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "<td valign=""top"" align=""left"" class=""availabledescription"">"

		response.write "<p><span class=""schedulerentalname"">"
		If oRs("locationname")  <> "" Then 
			response.write oRs("locationname") & " &ndash; " 
		End If 
		response.write oRs("rentalname")
		response.write "</span></p>"

		'response.write "<p>" & oRs("description") & "</p>"
		response.write "<p>"
		If oRs("usehtmlonlongdesc") Then 
			response.write oRs("description") 
		Else 
			response.write Replace(oRs("description"), Chr(10), "<br />")
		End If 
		response.write "</p>"

		If oRs("locationname")  <> "" Or oRs("width") <> "" Or oRs("capacity") <> "" Then 
			response.write vbcrlf & "<p>"
'			If oRs("locationname")  <> "" Then 
'				response.write "<strong>Location: </strong>" & oRs("locationname") & "<br />"
'			End If 
			If oRs("width") <> "" Then 
				response.write "<strong>Dimensions: </strong>" & oRs("width") & " x " & oRs("length") & "<br />"
			End If 
			If oRs("capacity") <> "" Then 
				response.write "<strong>Capacity: </strong>" & oRs("capacity") & "<br />"
			End If 
			response.write vbcrlf & "</p>"
		End If 

		DisplayRentalDocuments iRentalId

		response.write "</td>"
		response.write "</tr>"
		response.write "</table>"
		oRs.MoveNext 
	Loop

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' void SetTimeToOpeningAndClosing iRentalid, bOffSeasonFlag, iWeekday, dSelectedDate, iFirstHour, iFirstMinute, iFirstAmPm, iLastHour, iLastMinute, iLastAmPm
'--------------------------------------------------------------------------------------------------
Sub SetTimeToOpeningAndClosing( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dSelectedDate, ByRef iFirstHour, ByRef iFirstMinute, ByRef iFirstAmPm, ByRef iLastHour, ByRef iLastMinute, ByRef iLastAmPm )
	Dim dOpeningTime, dClosingTime

	GetOpeningAndClosingTimes iRentalid, bOffSeasonFlag, iWeekday, dSelectedDate, dOpeningTime, dClosingTime

	iFirstAmPm = "AM"
	iFirstHour = Hour(dOpeningTime)
	If clng(iFirstHour) = clng(0) Then
		iFirstHour = 12
		iFirstAmPm = "AM"
	Else
		If clng(iFirstHour) > clng(12) Then
			iFirstHour = clng(iFirstHour) - clng(12)
			iFirstAmPm = "PM"
		End If 
		If clng(iFirstHour) = clng(12) Then
			iFirstAmPm = "PM"
		End If 
	End If 
	iFirstMinute = Minute(dOpeningTime)
	If iFirstMinute < 10 Then
		iFirstMinute = "0" & iFirstMinute
	End If 

	iLastAmPm = "AM"
	iLastHour = Hour(dClosingTime)
	If clng(iLastHour) = clng(0) Then
		iLastHour = 12
		iLastAmPm = "AM"
	Else
		If clng(iLastHour) > clng(12) Then
			iLastHour = clng(iLastHour) - clng(12)
			iLastAmPm = "PM"
		End If 
		If clng(iLastHour) = clng(12) Then
			iLastAmPm = "PM"
		End If 
	End If 
	iLastMinute = Minute(dClosingTime)
	If iLastMinute < 10 Then
		iLastMinute = "0" & iLastMinute
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' void SetTimeToFirstAvailable iRentalid, bOffSeasonFlag, iWeekday, dSelectedDate, iFirstHour, iFirstMinute, iFirstAmPm, iLastHour, iLastMinute, iLastAmPm
'--------------------------------------------------------------------------------------------------
Sub SetTimeToFirstAvailable( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dSelectedDate, ByRef iFirstHour, ByRef iFirstMinute, ByRef iFirstAmPm, ByRef iLastHour, ByRef iLastMinute, ByRef iLastAmPm )
	Dim iInterval, sDateAddString, bHasMinimum, dFirstDateTime, dLastDateTime

	bHasMinimum = GetMinimalTimeInterval( iRentalid, bOffSeasonFlag, iWeekday, iInterval, sDateAddString )

	dFirstDateTime = GetFirstAvailableTime( iRentalId, bOffSeasonFlag, iWeekday, dSelectedDate )

	If bHasMinimum Then
		dLastDateTime = DateAdd(sDateAddString, iInterval, dFirstDateTime)
	Else
		dLastDateTime = DateAdd("h", 1, dFirstDateTime)
	End If 

	iFirstAmPm = "AM"
	iFirstHour = Hour(dFirstDateTime)
	If clng(iFirstHour) = clng(0) Then
		iFirstHour = 12
		iFirstAmPm = "AM"
	Else
		If clng(iFirstHour) > clng(12) Then
			iFirstHour = clng(iFirstHour) - clng(12)
			iFirstAmPm = "PM"
		End If 
		If clng(iFirstHour) = clng(12) Then
			iFirstAmPm = "PM"
		End If 
	End If 
	iFirstMinute = Minute(dFirstDateTime)
	If iFirstMinute < 10 Then
		iFirstMinute = "0" & iFirstMinute
	End If 

	iLastAmPm = "AM"
	iLastHour = Hour(dLastDateTime)
	If clng(iLastHour) = clng(0) Then
		iLastHour = 12
		iLastAmPm = "AM"
	Else
		If clng(iLastHour) > clng(12) Then
			iLastHour = clng(iLastHour) - clng(12)
			iLastAmPm = "PM"
		End If 
		If clng(iLastHour) = clng(12) Then
			iLastAmPm = "PM"
		End If 
	End If 
	iLastMinute = Minute(dLastDateTime)
	If iLastMinute < 10 Then
		iLastMinute = "0" & iLastMinute
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowOrgRoundUpTime iOrgId
'--------------------------------------------------------------------------------------------------
Sub ShowOrgRoundUpTime( ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(O.rentalroundup,0) AS rentalroundup, T.timetype "
	sSql = sSql & "FROM Organizations O, egov_rentaltimetypes T "
	sSql = sSql & "WHERE O.rentalrounduptimetypeid = T.timetypeid AND O.orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If CLng(oRs("rentalroundup")) > CLng(0) Then
			response.write "<br />*Reservation times will be automatically rounded up to the next " & oRs("rentalroundup") & " " 
			If LCase(oRs("timetype")) = "minutes" Then
				response.write "minute "
			Else  
				response.write "hour "
			End If 
			response.write "interval, if needed."
		End If 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRates iRentalid, bOffSeasonFlag, iWeekday, sResidentType
'--------------------------------------------------------------------------------------------------
Sub ShowRates( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal sResidentType )
	Dim sSql, oRs, sPriceTypeName, dBaseAmount, dAmount, sRateType, iBasePriceTypeId

	sSql = "SELECT R.pricetypeid, P.pricetypename, R.ratetypeid, ISNULL(amount,0.00) AS amount, T.ratetype, "
	sSql = sSql & "ISNULL(R.starthour,0) AS starthour, dbo.AddLeadingZeros(ISNULL(R.startminute,0),2) AS startminute, "
	sSql = sSql & "ISNULL(R.startampm,'AM') AS startampm, P.pricetype, P.isbaseprice, P.isfee, P.isweekendsurcharge, "
	sSql = sSql & "ISNULL(P.basepricetypeid,0) AS basepricetypeid, P.checkresidency, P.isresident, T.datediffstring "
	sSql = sSql & "FROM egov_rentaldayrates R, egov_rentaldays D, egov_price_types P, egov_rentalratetypes T "
	sSql = sSql & "WHERE D.dayid = R.dayid AND D.rentalid = R.rentalid AND R.pricetypeid = P.pricetypeid "
	sSql = sSql & "AND T.ratetypeid = R.ratetypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.isoffseason = " & bOffSeasonFlag & " AND D.dayofweek = " & iWeekday
	sSql = sSql & " ORDER BY P.displayorder"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		sPriceTypeName = oRs("pricetypename")
		sRateType = oRs("ratetype")
		If oRs("isbaseprice") Then 
			dBaseAmount = CDbl(oRs("amount"))
			dAmount = dBaseAmount
			iBasePriceTypeId = oRs("pricetypeid")
		Else
			If oRs("isfee") Then
				dAmount = CDbl(oRs("amount"))
				bShow = True
			Else 
				If CLng(iBasePriceTypeId) = CLng(oRs("basepricetypeid")) Then 
					dAmount = dBaseAmount + CDbl(oRs("amount"))	
				Else
					dAmount = CDbl(oRs("amount"))
				End If 
			End If 
		End If 
		If oRs("checkresidency") Then
			If oRs("pricetype") = sResidentType Then
				bShow = True 
			Else
				bShow = False 
			End If
		Else
			bShow = True 
		End If 

		If bShow Then 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""chargename"" valign=""center"">" & sPriceTypeName & "</td>"
			response.write "<td valign=""center"">" & FormatNumber(dAmount,2,,,0) & " " & LCase(sRateType)
			If clng(oRs("starthour")) > clng(0) Then 
				response.write " added for any time after " & oRs("starthour") & ":" & oRs("startminute") & " " & oRs("startampm")
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
' integer ShowRentalCharges( iRentalid, iIncludePriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ShowRentalCharges( ByVal iRentalid, ByVal iIncludePriceTypeId )
	Dim sSql, oRs, iCount

	iCount = 0

	' This pulls in the static fees like the deposit
	sSql = "SELECT P.pricetypeid, P.pricetypename, ISNULL(F.amount,0.00) AS amount, P.needsprompt, "
	sSql = sSql & " ISNULL(F.prompt,'') AS prompt "
	sSql = sSql & " FROM egov_price_types P, egov_rentalfees F "
	sSql = sSql & " WHERE F.pricetypeid = P.pricetypeid AND F.rentalid = " & iRentalid
	sSql = sSql & " ORDER BY P.displayorder"
	' response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iCount = iCount + 1
		response.write vbcrlf & "<tr>"
		response.write "<td class=""chargename"" valign=""center"">"
		response.write "<input type=""hidden"" name=""pricetypeid" & iCount & """ value=""" & oRs("pricetypeid") & """ />"
		response.write oRs("pricetypename") & "</td>"
		response.write "<td valign=""center"" nowrap=""nopwrap""><span class=""respCol"">" & FormatCurrency(oRs("amount"),2)  & "&nbsp;</span>"
		If oRs("needsprompt") Then
			response.write "<span class=""respCol""><input type=""checkbox"" name=""includepricetype" & iCount & """"
			If CLng(iIncludePriceTypeId) = CLng(oRs("pricetypeid")) Then
				response.write " checked=""checked"" "
			End If 
			response.write "/> <strong>" & oRs("prompt") & "</strong></span>"
		Else
			response.write "<input type=""hidden"" name=""includepricetype" & iCount & """ value=""off"" />"
		End If 
		response.write "</td>"
		response.write "</tr>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	ShowRentalCharges = iCount

End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalHasInSeasonOnly( iRentalId )
'--------------------------------------------------------------------------------------------------
Function RentalHasInSeasonOnly( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT reservationsduringseason FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("reservationsduringseason") Then 
			RentalHasInSeasonOnly = True 
		Else
			RentalHasInSeasonOnly = False 
		End If 
	Else
		RentalHasInSeasonOnly = False 
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean RentalHasNonResidentWait( iRentalId )
'--------------------------------------------------------------------------------------------------
Function RentalHasNonResidentWait( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT nonresidentswait, ISNULL(nonresidentwaitdays,0) AS nonresidentwaitdays "
	sSql = sSql & "FROM egov_rentals WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("nonresidentswait") And clng(oRs("nonresidentwaitdays")) > clng(0) Then 
			RentalHasNonResidentWait = True 
		Else
			RentalHasNonResidentWait = False 
		End If 
	Else
		RentalHasNonResidentWait = False 
	End If 

	oRs.Close 
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' date GetNonResidentStartDate( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetNonResidentStartDate( ByVal iRentalId )
	Dim bOffSeasonFlag, sSql, oRs, dSeasonStartDate, dSeasonEndDate

	bOffSeasonFlag = GetOffSeasonFlag( iRentalid, Date() )

	sSql = "SELECT hasoffseason, offseasonstartmonth, offseasonstartday, offseasonendmonth, offseasonendday, "
	sSql = sSql & "ISNULL(nonresidentwaitdays,0) AS nonresidentwaitdays "
	sSql = sSql & "FROM egov_rentals WHERE rentalid = " & iRentalid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If bOffSeasonFlag = "1" Then 
'			response.write "Off Season<Br />"
			' we are in an off season so set the start and end dates for that
			dSeasonStartDate = CDate(oRs("offseasonstartmonth") & "/" & oRs("offseasonstartday") & "/" & Year(Date()))
			'dSeasonEndDate = CDate(oRs("offseasonendmonth") & "/" & oRs("offseasonendday") & "/" & Year(Date()))
		Else
'			response.write "In Season<Br />"
			' current season start and end dates for currently being in season
			dSeasonStartDate = CDate(oRs("offseasonendmonth") & "/" & oRs("offseasonendday") & "/" & Year(Date()))
			'dSeasonEndDate = CDate(oRs("offseasonstartmonth") & "/" & oRs("offseasonstartday") & "/" & Year(Date()))
		End If 
		If dSeasonStartDate > Date() Then
			' we need to set the current start date to last year
			dSeasonStartDate = DateAdd("yyyy", -1, dSeasonStartDate)
		End If 
'		If dSeasonEndDate < Date() Then 
'			' we need to set the current end date to next year
'			dSeasonEndDate = DateAdd("yyyy", 1, dSeasonEndDate)
'		End If 
		GetNonResidentStartDate = DateAdd("d", clng(oRs("nonresidentwaitdays")), dSeasonStartDate )
	Else
		GetNonResidentStartDate = Date()
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean DateIsPastAllowedRange( iRentalId, sSelectedDate, sResidency, sLessThanDate )
'--------------------------------------------------------------------------------------------------
Function DateIsPastAllowedRange( ByVal iRentalId, ByVal sSelectedDate, ByVal sResidency, ByRef sLessThanDate )
	Dim sSql, oRs, sField, dEndRangeDate

	sLessThanDate = DateAdd( "d", 1, sSelectedDate )

	If sResidency = "R" Then
		sField = "residentrentalperiod"
	Else
		sField = "nonresidentrentalperiod"
	End If 

	sSql = "SELECT ISNULL( " & sField & ",9999) AS rentalperiod FROM egov_rentals WHERE rentalid = " & iRentalId
	sSql = sSql & " AND orgid = " & iOrgid
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		' the biggest number they can setup is 999
		If clng(oRs("rentalperiod")) = clng(9999) Then
			' They did not set one up for this rental, so any date goes
			DateIsPastAllowedRange = False
		Else
			sLessThanDate = DateAdd( "m", clng(oRs("rentalperiod")), DateValue(Now()) )
			If CDate(sLessThanDate) > CDate(sSelectedDate) Then
				' These are alowed dates in the range
				DateIsPastAllowedRange = False
			Else
				' These are too far out
				DateIsPastAllowedRange = True
			End If 
		End If 
	Else
		' There is no rental row, They have bigger problems ahead
		DateIsPastAllowedRange = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean NonresidentsCanPurchase( iRentalId, sSelectedDate )
'--------------------------------------------------------------------------------------------------
Function NonresidentsCanPurchase( ByVal iRentalId, ByVal sSelectedDayOfTheWeek )
	Dim sSql, oRs

	' To offer non-resident pricing, we need E or N type pricing set up'
	sSql = "SELECT COUNT(R.pricetypeid) AS offercount "
	sSql = sSql & "FROM egov_rentaldayrates R, egov_rentaldays D, egov_price_types P, egov_rentalratetypes T "
	sSql = sSql & "WHERE D.dayid = R.dayid AND D.rentalid = R.rentalid AND R.pricetypeid = P.pricetypeid "
	sSql = sSql & "AND T.ratetypeid = R.ratetypeid AND P.pricetype IN ('E','N') AND D.rentalid = " & iRentalId
	sSql = sSql & " AND D.isoffseason = 0 AND D.dayofweek = " & sSelectedDayOfTheWeek

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If clng(oRs("offercount")) > clng(0) Then
			NonresidentsCanPurchase = True
		Else
			NonresidentsCanPurchase = False	
		End If 
	Else
		NonresidentsCanPurchase = False
	End If

	oRs.Close
	Set oRs = Nothing 

End Function



%>

