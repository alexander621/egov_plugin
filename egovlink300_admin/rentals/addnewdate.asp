<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: addnewdate.asp
' AUTHOR: Steve Loar
' CREATED: 11/09/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:	Selects new dates and rentals to add to a reservation
'
' MODIFICATION HISTORY
' 1.0   11/09/2009	Steve Loar - INITIAL VERSION
' 1.1	05/11/2010	Steve Loar - Modified start and end times to be 12:00 AM
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iReservationId, iRentalId, sStartDate, sStartHour, sStartMinute, sStartAmPm, sEndHour
Dim sEndMinute, sEndAmPm, dStartDateTime, dEndDateTime, iEndDay, sLoadMsg, sMessage, bIsForAClass
Dim bOffSeasonFlag, bIsAllDayOnly, sDisabledOption
Dim aWantedDates(1,0)

iReservationId = CLng(request("reservationid"))
sLoadMsg = ""

bIsForAClass = ReservationIsForAClass( iReservationId )

If InStr(request("rentalid"), "R") > 0 Then
	iRentalId = CLng(Mid(request("rentalid"),2))
Else 
	iRentalId = CLng(request("rentalid"))
End If 

If request("startdate") = "" Then
	sStartDate = DateValue(Now())
Else
	sStartDate = DateValue(CDate(request("startdate")))
End If 

bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(sStartDate) )
If RentalIsAllDay( iRentalid, bOffSeasonFlag, Weekday(DateValue(sStartDate)) ) Then 
	bIsAllDayOnly = True 
	sDisabledOption = "disabled"
	dStartDateTime = sStartDate
	dEndDateTime = sStartDate
Else 
	bIsAllDayOnly = False 
	sDisabledOption = ""
End If 

' If this is for all day, or the rental is only available for all day reservations then we need the opening and closing times on that day
If bIsAllDayOnly Then
	SetHoursToOpenAndClose iRentalId, DateValue(sStartDate), dStartDateTime, dEndDateTime, iEndDay
	sStartHour = GetHourFromDateTime( dStartDateTime, sStartAmPm )  ' In rentalscommonfunctions.asp
	sStartMinute = Minute(dStartDateTime)
	sEndHour = GetHourFromDateTime( dEndDateTime, sEndAmPm )  ' In rentalscommonfunctions.asp
	sEndMinute = Minute(dEndDateTime)
Else 

	If request("starthour") = "" Then
		'sStartHour = GetHourFromDateTime( Now(), sStartAmPm )  ' In rentalscommonfunctions.asp
		sStartHour = "12"
		sStartAmPm = "AM"
	Else
		sStartHour = request("starthour")
		sStartAmPm = request("startampm")
	End If 

	If request("startminute") = "" Then
		'sStartMinute = Right(("00" & Minute(Now())), 2)
		sStartMinute = "00"
	Else
		sStartMinute = request("startminute")
	End If

	dStartDateTime = CDate((sStartDate & " " & sStartHour & ":" & sStartMinute & " " & sStartAmPm))

	If request("endday") = "" Then
		iEndDay = 0
		dEndDateTime = DateAdd("h", 1, dStartDateTime)
	Else
		iEndDay = request("endday")
		dEndDateTime = DateValue(DateAdd("d", iEndDay, dStartDateTime))
	End If 

	If request("endhour") = "" Then
		'sEndHour = GetHourFromDateTime( dEndDateTime, sEndAmPm )  ' In rentalscommonfunctions.asp
		sEndHour = "12"
		sEndAmPm = "AM"
	Else
		sEndHour = request("endhour")
		sEndAmPm = request("endampm")
	End If 

	If request("endminute") = "" Then
		'sEndMinute = Right(("00" & Minute(dEndDateTime)), 2)
		sEndMinute = "00"
	Else
		sEndMinute = request("endminute")
	End If

	dEndDateTime = CDate((DateValue(dEndDateTime) & " " & sEndHour & ":" & sEndMinute & " " & sEndAmPm))
End If 


aWantedDates(0,0) = dStartDateTime
aWantedDates(1,0) = dEndDateTime

' If this is a post then check availability and display any warnings.
If request.ServerVariables( "REQUEST_METHOD" ) = "POST" Then
	'response.write dStartDateTime & "<br />"
	'response.write dEndDateTime & "<br /><br />"

	' Round up as required by the org to the next wanted interval
	CheckOrgRentalRoundUp dStartDateTime, dEndDateTime, sEndHour, sEndMinute, sEndAmPm

	' Check the availability and get the message to show
	sMessage = CheckRentalAvailability( iRentalid, dStartDateTime, dEndDateTime, False )	' In rentalscommonfunctions.asp

	If sMessage = "short" Then
		If request("shortcheck") = "yes" Then 
			sLoadMsg = "displayScreenMsg('Warning: The duration is for less than the allowed minimum time.');"
		Else 
			sLoadMsg = "doShortConfirm();"
		End If 
	End If 

	If sMessage = "buffer" Then
		If request("shortcheck") = "yes" Then 
			sLoadMsg = "displayScreenMsg('Warning: There is a conflict with the buffering between reservations.');"
		Else
			sLoadMsg = "doBufferConfirm();"
		End If 
	End If
	If sMessage = "buffershort" Then
		If request("shortcheck") = "yes" Then 
			sLoadMsg = "displayScreenMsg('Warning: The duration is less than allowed and there is a conflict with the buffering.');"
		Else
			sLoadMsg = "doBufferShortConfirm();"
		End If 
	End If

	If sMessage = "conflict" Then
		sLoadMsg = "displayScreenMsg('There is a conflict with an existing reservation.');"
	End If 
	If sMessage = "closed" Then
		sLoadMsg = "displayScreenMsg('The rental is not open, or the time requested is beyond operating hours.');"
	End If 
	If sMessage = "OK" Then
		If request("reserve") = "yes" Then
			sLoadMsg = "sendOffToSave();"
		Else 
			sLoadMsg = "displayScreenMsg('Everything checks out fine for this reservation.');"
		End If 
	End If 
	'sLoadMsg = "displayScreenMsg('" & sMessage & "');"
End If 


%>

<html>
	<head>
		<link rel="stylesheet" type="text/css" href="../global.css" />
		<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

		<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

		<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
		<script language="javascript" src="../scripts/modules.js"></script>

		<script language="Javascript">
		<!--

		function doClose()
		{
			window.close();
			window.opener.focus();
		}

		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			var sSelectedDate = $(sField).value;

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmNewDate", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function displayScreenMsg( sMessage ) 
		{
			if(sMessage != "") 
			{
				$("screenMsg").innerHTML = "*** " + sMessage + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("screenMsg").innerHTML = "&nbsp;";
		}

		function CheckDates()
		{
			
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

			document.frmNewDate.submit();
		}

		function Validate()
		{
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

			// Set the flag so it behaves differently
			$("shortcheck").value = "no";
			$("reserve").value = "yes";

			document.frmNewDate.submit();
		}

		function loader()
		{
			// Display any messages or do other things
			<%=sLoadMsg%>
		}

		function doShortConfirm()
		{
			if (confirm("You have selected a time interval for one or more dates that is less than the allowed minimum.\nDo you wish to continue?"))
			{
				sendOffToSave();
				
			}
		}

		function doBufferConfirm()
		{
			if (confirm("There is a conflict with the buffering between reservations.\nDo you wish to continue?"))
			{
				sendOffToSave();
			}
		}

		function doBufferShortConfirm()
		{
			if (confirm("The duration is less than allowed and there is a conflict with the buffering.\nDo you wish to continue?"))
			{
				sendOffToSave();
			}
		}

		function sendOffToSave()
		{
			var sParameter = 'reservationid=' + encodeURIComponent($("reservationid").value);
			sParameter += '&rentalid=' + encodeURIComponent($("rentalid").value);
			sParameter += '&startdate=' + encodeURIComponent($("startdate").value);
			sParameter += '&starthour=' + encodeURIComponent($("starthour").value);
			sParameter += '&startminute=' + encodeURIComponent($("startminute").value);
			sParameter += '&startampm=' + encodeURIComponent($("startampm").value);
			sParameter += '&endhour=' + encodeURIComponent($("endhour").value);
			sParameter += '&endminute=' + encodeURIComponent($("endminute").value);
			sParameter += '&endampm=' + encodeURIComponent($("endampm").value);
			sParameter += '&endday=' + encodeURIComponent($("endday").value);
			//alert(sParameter);

			// Fire off job to save the date and times
			doAjax('addnewdatesave.asp', sParameter , 'saveReturn', 'post', '0');
		}

		function saveReturn( sReturn )
		{
			//alert( sReturn );
			window.opener.validateReservation();
			window.close();
			window.opener.focus();
		}

		//-->
		</script>

	</head>
	<body onload="loader();">
		<div id="content">
			<div id="centercontent">
				<p>
					<font size="+1"><strong>Add A Date</strong></font><br /><br />
				</p>
				<p>
					<span id="screenMsg">&nbsp;</span>
					&nbsp;
				</p>
				<form name="frmNewDate" action="addnewdate.asp" method="post">
					<input type="hidden" id="reservationid" name="reservationid" value="<%=iReservationId%>" />
					<input type="hidden" id="shortcheck" name="shortcheck" value="yes" />
					<input type="hidden" id="reserve" name="reserve" value="no" />
					<p>
					Rental: <% ShowRentalLocationPicks "R" & iRentalId, False, False 	' In rentalsguifunctions.asp %>
					</p><br />
					<p>
						<table id="reservationtempdates" cellpadding="0" cellspacing="0" border="0">
							<tr><th class="firstcell">Date</th><th>Start Time</th><th>End Time</th><th class="lastcell">Available</th></tr>
							<tr class="dateline">
								<td class="firstcell">
									<input type="text" id="startdate" name="startdate" value="<%=sStartDate%>" readonly="readonly" size="10" maxlength="10" onclick="javascript:void doCalendar('startdate');" />
									&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
								</td>
								<td align="center">
<%									ShowHourPicks "starthour", sStartHour, sDisabledOption %>:<% ShowMinutePicks "startminute", sStartMinute, sDisabledOption	%>&nbsp;
<%									ShowAmPmPicks "startampm", sStartAmPm, sDisabledOption		%>
								</td>
								<td align="center">
<%									ShowHourPicks "endhour", sEndHour, sDisabledOption %>:<% ShowMinutePicks "endminute", sEndMinute, sDisabledOption	%>&nbsp;
<%									ShowAmPmPicks "endampm", sEndAmPm, sDisabledOption		%>&nbsp;
<%									ShowSameNextDayPick "endday", iEndDay, sDisabledOption	%>
								</td>
								<td class="lastcell" align="center">
<%									ShowRentalAvailabilityFlag iRentalId, aWantedDates, "selectedperiod", bIsForAClass		%>
								</td>
							</tr>
							<tr>
								<td class="firstcell" colspan="2">
<%
								' Get rental details for that date
								response.write GetRentalSeason( iRentalId, DateValue(dStartDateTime) )
								response.write " &ndash; " & WeekDayName(Weekday(DateValue(dStartDateTime)))
								response.write GetRentalHours( iRentalId, DateValue(dStartDateTime) )
%>
								</td>
								<td class="lastcell" colspan="2" valign="top">Also happening on this date
<%									ShowOtherReservationsForDate iRentalId, DateValue(dStartDateTime)		%>
								</td>
							</tr>
							<tr><td colspan="4" class="tempseparator">&nbsp;</td></tr>
						</table>
					</p>
					<p>
						<input type="button" class="button" id="checkbutton" name="checkbutton" value="Check Dates" onclick="CheckDates()" />&nbsp; 
						<input type="button" class="button" id="continuebutton" name="continuebutton" value="Check and Reserve" onclick="Validate()" />&nbsp;
						<input type="button" class="button" value="Close" onclick="doClose();" /> 
					</p>
				</form>
			</div>
		</div>
	</body>
</html>

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------

%>