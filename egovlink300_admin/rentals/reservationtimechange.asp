<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationtimechange.asp
' AUTHOR: Steve Loar
' CREATED: 06/14/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Page that allows the time of a reservation to be changed on a specific date.
'
' MODIFICATION HISTORY
' 1.0   06/14/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationDateId, iReservationId, iRentalId, sReservationStartTime, sBillingEndTime, sLoadMsg
Dim sMessage, iReservationTypeId

If request("rdi") = "" Then 
	response.redirect "reservationlist.asp"
End If 

If Not IsNumeric(request("rdi")) Then
	response.redirect "reservationlist.asp"
End If 

sLevel = "../" ' Override of value from common.asp

sLoadMsg = ""

' check if page is online and user has permissions in one call not two
PageDisplayCheck "edit reservations", sLevel	' In common.asp

iReservationDateId = CLng(request("rdi"))

sMessage = request("msg")
If sMessage = "short" Then
	sLoadMsg = "doShortConfirm();"
End If 
If sMessage = "buffer" Then
	sLoadMsg = "doBufferConfirm();"
End If
If sMessage = "buffershort" Then
	sLoadMsg = "doBufferShortConfirm();"
End If
If sMessage = "shortnoconfirm" Then
	sLoadMsg = "displayScreenMsg('Warning: The duration is for less than the allowed minimum time.');"
End If 
If sMessage = "buffernoconfirm" Then
	sLoadMsg = "displayScreenMsg('Warning: There is a conflict with the buffering between reservations.');"
End If 
If sMessage = "buffershortnoconfirm" Then
	sLoadMsg = "displayScreenMsg('Warning: The duration is less than allowed and there is a conflict with the buffering.');"
End If 
If sMessage = "conflict" Then
	sLoadMsg = "displayScreenMsg('There is a conflict with an existing reservation.');"
End If 
If sMessage = "closed" Then
	sLoadMsg = "displayScreenMsg('The rental is not open, or the time requested is beyond operating hours.');"
End If 
If sMessage = "nouser" Then
	sLoadMsg = "displayScreenMsg('This type of reservation requires the selection of a person to complete.');"
End If 
If sMessage = "OK" Then
	sLoadMsg = "displayScreenMsg('The selected time checks out fine for this reservation.');"
End If 


GetReservationDateKeyValues iReservationDateId, iReservationId, iRentalId, sReservationStartTime, sBillingEndTime

iReservationTypeId = GetReservationTypeId( iReservationId )

If request("msg") <> "" Then
	' if this is filled then we are re-posting and we need to pick up the new start and end times
	sReservationStartTime = CDate(request("startdate") & " " & request("starthour") & ":" & request("startminute") & " " & request("startampm"))
	sBillingEndTime = CDate(request("startdate") & " " & request("endhour") & ":" & request("endminute") & " " & request("endampm"))
	If clng(request("endday")) > clng(0) Then
		sBillingEndTime = DateAdd("d", 1, sBillingEndTime)
	End If 
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	<!--

		function goBack()
		{
			document.frmTimeChange.action = 'reservationedit.asp';
			document.frmTimeChange.submit();
		}

		function loader()
		{
			<%=sLoadMsg%>
		}

		function doShortConfirm()
		{
			if (confirm("You have selected a time interval that is less than the allowed minimum.\nDo you wish to continue?"))
			{
				document.frmTimeChange.submit();
			}
		}
		
		function doBufferConfirm()
		{
			if (confirm("There is a conflict with the buffering between reservations.\nDo you wish to continue?"))
			{
				document.frmTimeChange.submit();
			}
		}

		function doBufferShortConfirm()
		{
			if (confirm("The duration is less than allowed and there is a conflict with the buffering.\nDo you wish to continue?"))
			{
				document.frmTimeChange.submit();
			}
		}

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "&nbsp;";
		}

		function CheckTimes()
		{

			// Check that the end time is later than the start times

			var dtStart = new Date($("startdate").value + " " + $("starthour").value + ":" + $("startminute").value + " " + $("startampm").value);
			var dtEnd;
			if ($("endday").value == "0")
			{
				dtEnd = new Date($("startdate").value + " " + $("endhour").value + ":" + $("endminute").value + " " + $("endampm").value);
				//alert(dtEnd);
			}
			else
			{
				dtEnd = new Date($("startdate").value + " " + $("endhour").value + ":" + $("endminute").value + " " + $("endampm").value);
				dtEnd.setDate(dtEnd.getDate()+1);
				//alert(dtEnd);
			}
			var difference_in_milliseconds = dtEnd - dtStart;
			if (difference_in_milliseconds <= 0)
			{
				alert("The end time is not after the start time. Please correct this and try again.");
				$("endhour").focus();
				return;
			}

			//alert("OK");
			//return;
			// Now bundle the times and send off to check routine via AJAX
			var sParameter = 'rentalid=' + encodeURIComponent($("rentalid").value);
			sParameter += '&reservationdateid=' + encodeURIComponent($("rdi").value);
			sParameter += '&reservationid=' + encodeURIComponent($("reservationid").value);
			sParameter += '&startdate=' + encodeURIComponent($("startdate").value);
			sParameter += '&starthour=' + encodeURIComponent($("starthour").value);
			sParameter += '&startminute=' + encodeURIComponent($("startminute").value);
			sParameter += '&startampm=' + encodeURIComponent($("startampm").value);
			sParameter += '&endhour=' + encodeURIComponent($("endhour").value);
			sParameter += '&endminute=' + encodeURIComponent($("endminute").value);
			sParameter += '&endampm=' + encodeURIComponent($("endampm").value);
			sParameter += '&endday=' + encodeURIComponent($("endday").value);

			// Fire off job to check times
			doAjax('checkselectedtimes.asp', sParameter , 'checkReturn', 'post', '0');
		}

		function checkReturn( sReturn )
		{
			//alert(sReturn);
			//return;

			document.frmTimeChange.action = "reservationtimechange.asp";
			if (sReturn != 'short' && sReturn != 'buffer' && sReturn != 'buffershort')
			{
				document.frmTimeChange.msg.value = sReturn;
			}
			else
			{
				document.frmTimeChange.msg.value = sReturn + 'noconfirm';
			}
			document.frmTimeChange.submit();
		}


		function Validate()
		{

			// Check that the end time is later than the start times
			var dtStart = new Date($("startdate").value + " " + $("starthour").value + ":" + $("startminute").value + " " + $("startampm").value);
			var dtEnd;
			if ($("endday").value == "0")
			{
				dtEnd = new Date($("startdate").value + " " + $("endhour").value + ":" + $("endminute").value + " " + $("endampm").value);
				//alert(dtEnd);
			}
			else
			{
				dtEnd = new Date($("startdate").value + " " + $("endhour").value + ":" + $("endminute").value + " " + $("endampm").value);
				dtEnd.setDate(dtEnd.getDate()+1);
				//alert(dtEnd);
			}
			var difference_in_milliseconds = dtEnd - dtStart;
			if (difference_in_milliseconds <= 0)
			{
				alert("The end time is not after the start time. Please correct this and try again.");
				$("endhour").focus();
				return;
			}

			//alert("OK");
			//return;
			// Now bundle the dates and times and send off to check routine via AJAX
			var sParameter = 'rentalid=' + encodeURIComponent($("rentalid").value);
			sParameter += '&reservationdateid=' + encodeURIComponent($("rdi").value);
			sParameter += '&reservationid=' + encodeURIComponent($("reservationid").value);
			sParameter += '&startdate=' + encodeURIComponent($("startdate").value);
			sParameter += '&starthour=' + encodeURIComponent($("starthour").value);
			sParameter += '&startminute=' + encodeURIComponent($("startminute").value);
			sParameter += '&startampm=' + encodeURIComponent($("startampm").value);
			sParameter += '&endhour=' + encodeURIComponent($("endhour").value);
			sParameter += '&endminute=' + encodeURIComponent($("endminute").value);
			sParameter += '&endampm=' + encodeURIComponent($("endampm").value);
			sParameter += '&endday=' + encodeURIComponent($("endday").value);

			// Fire off job to check dates and times
			doAjax('checkselectedtimes.asp', sParameter , 'validateReturn', 'post', '0');
		}

		function validateReturn( sReturn )
		{
			//alert(sReturn);
			if (sReturn != 'OK')
			{
				document.frmTimeChange.action = "reservationtimechange.asp";
				document.frmTimeChange.msg.value = sReturn;
				document.frmTimeChange.submit();
			}
			else
			{
				document.frmTimeChange.submit();
			}
		}

		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			var sSelectedDate = $(sField).value;

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmTimeChange", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

	//-->
	</script>

</head>

<body onload="loader();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Reservation Time Change</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<p>
			<span id="screenMsg">&nbsp;</span>
			&nbsp;
			</p>

			<p>
				<input type="button" class="button" id="back" name="back" value="<< Back to Reservation" onclick="goBack();" /> &nbsp;
			</p>

			<form name="frmTimeChange" method="post" action="reservationtimechangedo.asp">
				<input type="hidden" id="rdi" name="rdi" value="<%=iReservationDateId%>" />
				<input type="hidden" id="reservationid" name="reservationid" value="<%=iReservationId%>" />
				<input type="hidden" id="rentalid" name="rentalid" value="<%=iRentalId%>" />
				<input type="hidden" id="msg" name="msg" value="<%=sMessage%>" />

				<% ShowRentalNameAndLocation iRentalId %><br /><br />
				<p>
					<table id="reservationtempdates" cellpadding="0" cellspacing="0" border="0">
						<tr><th class="firstcell">Date</th><th>Start Time</th><th>End Time</th><th class="lastcell">Available</th></tr>
<%						
						ShowReservationDateAndTimes iReservationDateId, iRentalId, sReservationStartTime, sBillingEndTime, iReservationTypeId
%>			
					</table>
				</p>
				<p id="usagenote">
					Note: You can change the time of the reservation for the selected date only.
				</p>
				<p>
					<input type="button" class="button" id="checkbutton" name="checkbutton" value="Check Times" onclick="CheckTimes()" />&nbsp;
					<input type="button" class="button" id="continuebutton" name="continuebutton" value="Check and Reserve" onclick="Validate()" />
					<input type="hidden" id="oldstartdate" name="oldstartdate" value="<%=DateValue(sReservationStartTime)%>" />
				</p>

			</form>


		</div>
	</div>

	<!--END: PAGE CONTENT-->

	<!--#Include file="../admin_footer.asp"-->  

</body>

</html>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void GetReservationDateKeyValues iReservationDateId, iReservationId, sReservationStartTime, sBillingEndTime
'--------------------------------------------------------------------------------------------------
Sub GetReservationDateKeyValues( ByVal iReservationDateId, ByRef iReservationId, ByRef iRentalId, ByRef sReservationStartTime, ByRef sBillingEndTime )
	Dim sSql, oRs

	sSql = "SELECT reservationid, rentalid, reservationstarttime, billingendtime FROM egov_rentalreservationdates "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND reservationdateid = " & iReservationDateId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iReservationId = oRs("reservationid")
		iRentalId = oRs("rentalid")
		sReservationStartTime = oRs("reservationstarttime")
		sBillingEndTime = oRs("billingendtime")
		'session("sReservationStartTime") = sReservationStartTime
		'session("sBillingEndTime") = sBillingEndTime
	Else
		iReservationId = 0
		iRentalId = 0
		sReservationStartTime = Now()
		sBillingEndTime = Now()
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReservationDateAndTimes iReservationDateId, iRentalId, sReservationStartTime, sBillingEndTime, iReservationTypeId
'--------------------------------------------------------------------------------------------------
Sub ShowReservationDateAndTimes( ByVal iReservationDateId, ByVal iRentalId, ByVal sReservationStartTime, ByVal sBillingEndTime, ByVal iReservationTypeId )
	Dim bOffSeasonFlag, bIsAllDayOnly, sDisabledOption, sAmPm, iEndDay

	iEndDay = Abs(DateDiff("d", sReservationStartTime, sBillingEndTime ))

	' Build the date time row
	response.write vbcrlf & "<tr class=""dateline"">"
		 
	response.write "<td class=""firstcell"" align=""center"">"
'	response.write "<span class=""reservationdata"">"
	'response.write DateValue(sReservationStartTime)
'	response.write "</span>"
	response.write "<input type=""text"" id=""startdate"" name=""startdate"" value=""" & DateValue(sReservationStartTime) & """ readonly=""readonly"" size=""10"" maxlength=""10"" onclick=""javascript:void doCalendar('startdate');"" />"
	response.write "&nbsp;<span class=""calendarimg""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('startdate');"" /></span>"

	response.write "</td>"

	bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(sReservationStartTime) )
	If RentalIsAllDay( iRentalid, bOffSeasonFlag, Weekday(DateValue(sReservationStartTime)) ) Then 
		bIsAllDayOnly = True 
		sDisabledOption = "disabled"
		sPeriodTypeSelector = "allday"
	Else 
		bIsAllDayOnly = False 
		sDisabledOption = ""
		sPeriodTypeSelector = "selectedperiod"
		' if this is not a class
		If IsReservation( iReservationTypeId ) Then 
		' Round up as required by the org to the next wanted interval
			CheckOrgRentalRoundUp sReservationStartTime, sBillingEndTime, iEndHour, iEndMinute, sEndAmPm
			'session("sBillingEndTime") = "round" & sBillingEndTime
		Else
			SetEndingTimes sBillingEndTime, iEndHour, iEndMinute, sEndAmPm 
			'session("sBillingEndTime") = "set" & sBillingEndTime
		End If 
		'session("sBillingEndTime") = ""
		sBillingEndTime = DateValue(sBillingEndTime) & " " & iEndHour & ":" & Right("0" & iEndMinute,2) & " " & sEndAmPm
		'session("sBillingEndTime") = sBillingEndTime
		'response.write "sBillingEndTime = " & sBillingEndTime & "<br /><br />"
	End If 
	'response.write "sPeriodTypeSelector = " & sPeriodTypeSelector

	' show the start time picks
	response.write "<td align=""center"">" 
	'session("sReservationStartTime") = sReservationStartTime
	ShowHourPicks "starthour", GetHourFromDateTime( sReservationStartTime, sAmPm ), sDisabledOption  ' In rentalsguifunctions.asp
	response.write ":"
	ShowMinutePicks "startminute", Minute(sReservationStartTime), sDisabledOption	  ' In rentalsguifunctions.asp
	response.write " "
	ShowAmPmPicks "startampm", sAmPm, sDisabledOption	  ' In rentalsguifunctions.asp
	response.write "</td>"

	' Show the End time picks
	response.write "<td align=""center"">" 
	'session("sBillingEndTime") = sBillingEndTime
	ShowHourPicks "endhour", GetHourFromDateTime( sBillingEndTime, sAmPm ), sDisabledOption	  ' In rentalsguifunctions.asp
	response.write ":"
	ShowMinutePicks "endminute", Minute(sBillingEndTime), sDisabledOption	  ' In rentalsguifunctions.asp
	response.write " "
	ShowAmPmPicks "endampm", sAmPm, sDisabledOption	  ' In rentalsguifunctions.asp
	response.write " "
	ShowSameNextDayPick "endday", iEndDay, sDisabledOption	  ' In rentalsguifunctions.asp
	response.write "</td>"

	' Get the availability flag on that date and time
	response.write "<td class=""lastcell"" align=""center"">"
	ShowTimeAvailabilityFlag iRentalId, iReservationDateId, sReservationStartTime, sBillingEndTime, bOffSeasonFlag, sPeriodTypeSelector, iReservationId
	response.write "</td>"
	
	response.write "</tr>"

	response.write vbcrlf & "<tr>"
	response.write "<td class=""firstcell tablebottom"" colspan=""2"">"
	' Get rental details for that date
	response.write WeekDayName(Weekday(DateValue(sReservationStartTime)))
	response.write " &ndash; " & GetRentalSeason( iRentalId, DateValue(sReservationStartTime) )
	response.write GetRentalHours( iRentalId, DateValue(sReservationStartTime) )
	response.write "</td>"

	response.write "<td class=""lastcell tablebottom"" colspan=""2"" valign=""top"">Also happening on this date"
	' Get the other reservations, etc for this date
	ShowOtherReservationsOnThisDate iRentalId, DateValue(sReservationStartTime), iReservationDateId
	response.write "</td>"
	response.write "</tr>"

	' The seperator Row
	'response.write "<tr><td colspan=""4"" class=""tempseparator"">&nbsp;</td></tr>"


End Sub 


'--------------------------------------------------------------------------------------------------
' ShowTimeAvailabilityFlag iRentalId, iReservationDateId, sReservationStartTime, sBillingEndTime, bOffSeasonFlag, sPeriodTypeSelector
'--------------------------------------------------------------------------------------------------
Sub ShowTimeAvailabilityFlag( ByVal iRentalId, ByVal iReservationDateId, ByVal sReservationStartTime, ByVal sBillingEndTime, ByVal bOffSeasonFlag, ByVal sPeriodTypeSelector, ByVal iReservationId )

	' Now check that the wanted date fits into the hours of the rental itself
	sFlag = CheckRentalHours( iRentalid, sReservationStartTime, sBillingEndTime, sPeriodTypeSelector, bOffSeasonFlag )

	If UCase(sFlag) = "YES" Then
		sFlag = CheckIfTimeIsAvailable( iRentalid, sReservationStartTime, sBillingEndTime, iReservationDateId, sPeriodTypeSelector, iReservationId, bOffSeasonFlag )
	End If 

	response.write sFlag

End Sub


'--------------------------------------------------------------------------------------------------
' String CheckIfTimeIsAvailable( iRentalid, dWantedStartTime, dWantedEndTime, iReservationDateId, sPeriodTypeSelector, bOffSeasonFlag )
'--------------------------------------------------------------------------------------------------
Function CheckIfTimeIsAvailable( ByVal iRentalid, ByVal dWantedStartTime, ByVal dWantedEndTime, ByVal iReservationDateId, ByVal sPeriodTypeSelector, ByVal iReservationId, ByVal bOffSeasonFlag )
	Dim sCompareEndTime, bIncludeBuffer

	If sPeriodTypeSelector = "selectedperiod" Then
		bIncludeBuffer = ReservationNeedsBufferTimeAdded( iReservationId )

		If bIncludeBuffer Then 
			' Add on the end buffer
			dWantedEndTime = AddPostBufferTime( iRentalid, bOffSeasonFlag, dWantedEndTime, dWantedStartTime )
		End If 

		If bIncludeBuffer Then
			sCompareEndTime = "reservationendtime"
		Else
			sCompareEndTime = "billingendtime"
		End If 

		' we will add a minute to this so start time can be the same as the end of bufferend time of another reservation
		dWantedStartTime = DateAdd("n", 1, dWantedStartTime)
		' we will remove a minute so the end of the buffer can be the same minute as the start of another reservation
		dWantedEndTime = DateAdd("n", -1, dWantedEndTime)

		' set sql to look for conflicting times
		sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
		sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
		sSql = sSql & " AND (reservationstarttime BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
		sSql = sSql & " OR " & sCompareEndTime & " BETWEEN '" & dWantedStartTime & "' AND '" & dWantedEndTime & "' "
		sSql = sSql & " OR (reservationstarttime <= '" & dWantedStartTime & "' AND " & sCompareEndTime & " >= '" & dWantedEndTime & "'))"
		sSql = sSql & " AND reservationdateid != " & iReservationDateId
	Else
		' allday - set sql to look for any starting time on that day
		sSql = "SELECT COUNT(reservationdateid) AS hits FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
		sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
		sSql = sSql & " AND reservationstarttime > '" & DateValue(dWantedStartTime) & " 0:00 AM' "
		sSql = sSql & " AND reservationstarttime < '" & DateValue(DateAdd("d", 1, dWantedStartTime)) & " 0:00 AM' "
		sSql = sSql & " AND reservationdateid != " & iReservationDateId
	End If 
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			' conflicting reservation times
			CheckIfTimeIsAvailable = "No"
		Else
			' No conflicts found
			CheckIfTimeIsAvailable = "Yes"
		End If 
	Else
		' No rows returned - not likely using count(), but still no conflicts
		CheckIfTimeIsAvailable = "Yes"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowOtherReservationsOnThisDate iRentalId, dReservationStartTime, iReservationDateId
'--------------------------------------------------------------------------------------------------
Sub ShowOtherReservationsOnThisDate( ByVal iRentalId, ByVal dReservationStartTime, ByVal iReservationDateId )
	Dim oRs, sSql, dWantedEndTime, sFirstName, sLastName, sPhone, dStartDateTime

	' This gets anything that starts anytime on the passed date
	dStartDateTime = dReservationStartTime & " 0:00 AM" ' Add the time of midnight to the passed in date
	dWantedEndTime = DateAdd("d", 1, CDate(dReservationStartTime)) ' set this to midnight of the next day

	sSql = "SELECT D.reservationid, D.reservationstarttime, D.billingendtime, D.reservationendtime, T.reservationtype, R.timeid, T.reservationtypeselector, "
	sSql = sSql & " ISNULL(R.rentaluserid,0) AS rentaluserid, ISNULL(R.adminuserid,'') AS adminuserid, T.isreservation, T.isclass "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservations R, egov_rentalreservationtypes T "
	sSql = sSql & " WHERE D.reservationid = R.reservationid AND R.reservationtypeid = T.reservationtypeid AND D.rentalid = " & iRentalid
	sSql = sSql & " AND D.statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
	sSql = sSql & " AND D.reservationstarttime BETWEEN '" & dStartDateTime & "' AND '" & dWantedEndTime & "' "
	sSql = sSql & " AND reservationdateid != " & iReservationDateId
	sSql = sSql & " ORDER BY D.reservationstarttime"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		' GetTimeFormat is in common.asp
		response.write "<br /><a href=""reservationedit.asp?reservationid=" & oRs("reservationid") & """ target=""_blank"">"
		response.write GetTimeFormat(oRs("reservationstarttime")) & " to " & GetTimeFormat(oRs("billingendtime")) & " &ndash; "
		response.write oRs("reservationtype") & " &ndash; "
		If oRs("isreservation") Then
			If oRs("reservationtypeselector") = "public" Then 
				' Show the citizen name
				ShowShortCitizenName oRs("rentaluserid")
			Else
				response.write GetAdminName(oRs("rentaluserid"))
			End If 
		Else
			If oRs("isclass") Then
				' Show the class name
				ShowShortClassName oRs("timeid")
			Else 
				' the admin who made the hold or block or whatever - in rentlascommonfunctions.asp
				GetAdminNameAndPhone oRs("adminuserid"), sFirstName, sLastName, sPhone
				If sLastName <> "" And sFirstName <> "" Then 
					response.write Left(UCase(Left(sFirstName,1)) & ". " & sLastName,30)
				End If 
			End If 
		End If 
		response.write "</a>"
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
