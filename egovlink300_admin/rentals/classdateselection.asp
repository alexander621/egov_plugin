<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: classdateselection.asp
' AUTHOR: Steve Loar
' CREATED: 12/8/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Class date and time selection page in the reservation process.
'
' MODIFICATION HISTORY
' 1.0   12/8/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iRentalId, iTimeId, iReservationTypeId, sReservationType, sClassName, sStartDate, sEndDate
Dim sRentalName, sLocationName, iClassId, sMessage, iReservationTempId, bIsNewRTI

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "make reservations", sLevel	' In common.asp

iRentalId = CLng(request("rentalid"))

iTimeId = CLng(request("timeid"))

' Set the rentalid for this time id back on the class so it shows the one selected when they came here
SetRentalIdForTimeId iTimeId, iRentalId

iReservationTypeId = GetReservationTypeIdBySelector( "class" )

sReservationType = GetReservationType( iReservationTypeId )

GetClassInformation iTimeId, iClassId, sClassName, sStartDate, sEndDate 

If request("rti") <> "" Then
	iReservationTempId = CLng(request("rti"))
	bIsNewRTI = False 
Else
	iReservationTempId = SaveReservationTempInfo( iRentalId, iTimeId )
	bIsNewRTI = True 
End If 

sMessage = request("msg")
If sMessage = "conflict" Then
	sLoadMsg = "displayScreenMsg('There is a conflict with an existing reservation.');"
End If 
If sMessage = "closed" Then
	sLoadMsg = "displayScreenMsg('The rental is not open, or the time requested is beyond the time restrictions.');"
End If 
If sMessage = "OK" Then
	sLoadMsg = "displayScreenMsg('Everything checks out fine for this reservation.');"
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

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmDateSelection", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function goBack()
		{
			location.href='../classes/edit_class.asp?classid=<%=iClassId%>';
		}

		function loader()
		{
			<%=sLoadMsg%>
		}

		function CheckDates()
		{
			for (var i = 1; i <= parseInt($("maxrows").value); i++)
			{
				// See if a row exists for this one
				if ($("startdate" + i))
				{
					if ($F("startdate" + i) == '')
					{
						alert("Please enter a date, then try again.");
						$("startdate" + i).focus();
						return;
					}
				}
			}

			// Check that all end times are later than start times
			for (var i = 1; i <= parseInt($("maxrows").value); i++)
			{
				// See if a row exists for this one
				if ($("startdate" + i))
				{
					var dtStart = new Date($("startdate" + i).value + " " + $("starthour" + i).value + ":" + $("startminute" + i).value + " " + $("startampm" + i).value);
					var dtEnd;
					if ($("endday" + i).value == "0")
					{
						dtEnd = new Date($("startdate" + i).value + " " + $("endhour" + i).value + ":" + $("endminute" + i).value + " " + $("endampm" + i).value);
						//alert(dtEnd);
					}
					else
					{
						dtEnd = new Date($("startdate" + i).value + " " + $("endhour" + i).value + ":" + $("endminute" + i).value + " " + $("endampm" + i).value);
						dtEnd.setDate(dtEnd.getDate()+1);
						//alert(dtEnd);
					}
					var difference_in_milliseconds = dtEnd - dtStart;
					if (difference_in_milliseconds <= 0)
					{
						alert("One of the end times is not after the start time. Please correct this and try again.");
						$("endhour" + i).focus();
						return;
					}
				}
			}
			//alert("OK");
			//return;
			// Now bundle the dates and times and send off to check routine via AJAX
			var sParameter = 'rentalid=' + encodeURIComponent($("rentalid").value);
			sParameter += '&maxrows=' + encodeURIComponent($("maxrows").value);
			sParameter += '&timeid=' + encodeURIComponent($("timeid").value);
			sParameter += '&rti=' + encodeURIComponent($("rti").value);
			for (var t = 1; t <= parseInt($("maxrows").value); t++)
			{
				if ($("startdate" + t))
				{
					sParameter += '&includereservationtime' + t + '=' + encodeURIComponent($("includereservationtime" + t).checked);
					sParameter += '&startdate' + t + '=' + encodeURIComponent($("startdate" + t).value);
					sParameter += '&starthour' + t + '=' + encodeURIComponent($("starthour" + t).value);
					sParameter += '&startminute' + t + '=' + encodeURIComponent($("startminute" + t).value);
					sParameter += '&startampm' + t + '=' + encodeURIComponent($("startampm" + t).value);
					sParameter += '&endhour' + t + '=' + encodeURIComponent($("endhour" + t).value);
					sParameter += '&endminute' + t + '=' + encodeURIComponent($("endminute" + t).value);
					sParameter += '&endampm' + t + '=' + encodeURIComponent($("endampm" + t).value);
					sParameter += '&endday' + t + '=' + encodeURIComponent($("endday" + t).value);
					sParameter += '&timedayid' + t + '=' + encodeURIComponent($("timedayid" + t).value);
				}
			}

			// Fire off job to check dates and times
			doAjax('checkselectedclassdates.asp', sParameter , 'checkReturn', 'post', '0');
		}

		function checkReturn( sReturn )
		{
			//alert(sReturn);
			document.frmDateSelection.action = "classdateselection.asp";
			document.frmDateSelection.msg.value = sReturn;
			document.frmDateSelection.submit();
		}

		function Validate()
		{
			for (var i = 1; i <= parseInt($("maxrows").value); i++)
			{
				// See if a row exists for this one
				if ($("startdate" + i))
				{
					if ($F("startdate" + i) == '')
					{
						alert("Please enter a date, then try again.");
						$("startdate" + i).focus();
						return;
					}
				}
			}

			// Check that all end times are later than start times
			for (var i = 1; i <= parseInt($("maxrows").value); i++)
			{
				// See if a row exists for this one
				if ($("startdate" + i))
				{
					var dtStart = new Date($("startdate" + i).value + " " + $("starthour" + i).value + ":" + $("startminute" + i).value + " " + $("startampm" + i).value);
					var dtEnd;
					if ($("endday" + i).value == "0")
					{
						dtEnd = new Date($("startdate" + i).value + " " + $("endhour" + i).value + ":" + $("endminute" + i).value + " " + $("endampm" + i).value);
						//alert(dtEnd);
					}
					else
					{
						dtEnd = new Date($("startdate" + i).value + " " + $("endhour" + i).value + ":" + $("endminute" + i).value + " " + $("endampm" + i).value);
						dtEnd.setDate(dtEnd.getDate()+1);
						//alert(dtEnd);
					}
					var difference_in_milliseconds = dtEnd - dtStart;
					if (difference_in_milliseconds <= 0)
					{
						alert("One of the end times is not after the start time. Please correct this and try again.");
						$("endhour" + i).focus();
						return;
					}
				}
			}
			//alert("OK");
			//return;
			// Now bundle the dates and times and send off to check routine via AJAX
			var sParameter = 'rentalid=' + encodeURIComponent($("rentalid").value);
			sParameter += '&maxrows=' + encodeURIComponent($("maxrows").value);
			sParameter += '&timeid=' + encodeURIComponent($("timeid").value);
			sParameter += '&rti=' + encodeURIComponent($("rti").value);
			for (var t = 1; t <= parseInt($("maxrows").value); t++)
			{
				if ($("startdate" + t))
				{
					sParameter += '&includereservationtime' + t + '=' + encodeURIComponent($("includereservationtime" + t).checked);
					sParameter += '&startdate' + t + '=' + encodeURIComponent($("startdate" + t).value);
					sParameter += '&starthour' + t + '=' + encodeURIComponent($("starthour" + t).value);
					sParameter += '&startminute' + t + '=' + encodeURIComponent($("startminute" + t).value);
					sParameter += '&startampm' + t + '=' + encodeURIComponent($("startampm" + t).value);
					sParameter += '&endhour' + t + '=' + encodeURIComponent($("endhour" + t).value);
					sParameter += '&endminute' + t + '=' + encodeURIComponent($("endminute" + t).value);
					sParameter += '&endampm' + t + '=' + encodeURIComponent($("endampm" + t).value);
					sParameter += '&endday' + t + '=' + encodeURIComponent($("endday" + t).value);
					sParameter += '&timedayid' + t + '=' + encodeURIComponent($("timedayid" + t).value);
				}
			}

			// Fire off job to check dates and times
			doAjax('checkselectedclassdates.asp', sParameter , 'validateReturn', 'post', '0');
		}

		function validateReturn( sReturn )
		{
			if (sReturn != 'OK')
			{
				document.frmDateSelection.action = "classdateselection.asp";
				document.frmDateSelection.msg.value = sReturn;
			}

			document.frmDateSelection.submit();
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
				<font size="+1"><strong>Class Rental Date Selection</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<p>
			<span id="screenMsg">&nbsp;</span>
			&nbsp;
			</p>
			
			<form name="frmDateSelection" method="post" action="classreservationmake.asp">
				<input type="hidden" id="rentalid" name="rentalid" value="<%=iRentalId%>" />
				<input type="hidden" id="timeid" name="timeid" value="<%=iTimeId%>" />
				<input type="hidden" id="rti" name="rti" value="<%=iReservationTempId%>" />
				<input type="hidden" name="msg" value="<%=sMessage%>" />
				<input type="hidden" id="reservationtypeid" name="reservationtypeid" value="<%=iReservationTypeId%>" />

				<% ShowRentalNameAndLocation iRentalId %><br /><br />

				<table id="reservationclassinfo" cellpadding="0" cellspacing="1" border="0">
				<tr>
					<td class="labelcolumn"><strong>Reservation Type:</strong></td>
					<td class="datacolumn" align="left" colspan="3"><%=sReservationType%></td>
				</tr>
				<tr>
					<td class="labelcolumn"><strong>Class:</strong></td>
					<td class="datacolumn" align="left" colspan="3"><%=sClassName%></td>
				</tr>
				<tr>
					<td class="labelcolumn"><strong>Start Date:</strong></td>
					<td class="datacolumn" align="left" colspan="3"><%=sStartDate%></td>
				</tr>
				<tr>
					<td class="labelcolumn"><strong>End Date:</strong></td>
					<td class="datacolumn" align="left" colspan="3"><%=sEndDate%></td>
				</tr>
				</table>

				<p>
					<input type="button" class="button" id="back" name="back" value="<< Back" onclick="goBack();" />
				</p>

<%				ShowRentalAvailabilityDetailsByActivityDays iTimeId, iRentalId, sStartDate, sEndDate, bIsNewRTI	%>

			</form>

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
' GetClassInformation iTimeId, sClassName, sStartDate, sEndDate 
'--------------------------------------------------------------------------------------------------
Sub GetClassInformation( ByVal iTimeId, ByRef iClassId, ByRef sClassName, ByRef sStartDate, ByRef sEndDate )
	Dim sSql, oRs

	sSql = "SELECT C.classid, C.classname, C.startdate, C.enddate "
	sSql = sSql & "FROM egov_class C, egov_class_time T "
	sSql = sSql & "WHERE C.classid = T.classid AND T.timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iClassId = oRs("classid")
		sClassName = oRs("classname")
		sStartDate = DateValue(oRs("startdate"))
		sEndDate = DateValue(oRs("enddate"))
	Else
		iClassId = 0
		sClassName = ""
		sStartDate = ""
		sEndDate = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' GetRentalInformation iRentalId, sRentalName, sLocationName
'--------------------------------------------------------------------------------------------------
Sub GetRentalInformation( ByVal iRentalId, ByRef sRentalName, ByRef sLocationName )
	Dim sSql, oRs

	sSql = "SELECT R.rentalname, L.name AS locationname "
	sSql = sSql & "FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.locationid = L.locationid AND R.rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sRentalName = oRs("rentalname")
		sLocationName = oRs("locationname")
	Else
		sRentalName = ""
		sLocationName = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' SetRentalIdForTimeId iTimeId, iRentalId
'--------------------------------------------------------------------------------------------------
Sub SetRentalIdForTimeId( ByVal iTimeId, ByVal iRentalId )
	Dim sSql

	sSql = "UPDATE egov_class_time SET rentalid = " & iRentalId
	sSql = sSql & " WHERE timeid = " & iTimeId

	RunSQLStatement sSql

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowRentalAvailabilityDetailsByActivityDays iTimeId, iRentalId, sStartDate, sEndDate, bIsNewRTI
'--------------------------------------------------------------------------------------------------
Sub ShowRentalAvailabilityDetailsByActivityDays( ByVal iTimeId, ByVal iRentalId, ByVal sStartDate, ByVal sEndDate, ByVal bIsNewRTI )
	Dim sSql, oRs, sDayCell, bHaveDay, sWantedDOWs, sStartTime, sEndTime, iTotalRows
	Dim aWantedDates()

	sDayCell = ""
	bHaveDay = False
	sWantedDOWs = ""
	iTotalRows = 0
	iTableCount = 0

	sSql = "SELECT D.timedayid, T.activityno, D.starttime, D.endtime, sunday, monday, tuesday, wednesday, thursday, friday, saturday "
	sSql = sSql & " FROM egov_class_time T, egov_class_time_days D "
	sSql = sSql & " WHERE T.timeid = D.timeid AND T.timeid = " & iTimeId
	sSql = sSql & " ORDER BY D.timedayid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		sDayCell = ""
		bHaveDay = False
		sWantedDOWs = ""
		response.write vbcrlf & "<p class=""activitydisplay"">Activity No: " & oRs("activityno")
		response.write "<span class=""timedisplay"">" & oRs("starttime") & " &ndash; " & oRs("endtime") & "</span>"
		If oRs("sunday") Then 
			sDayCell = "Sunday"
			bHaveDay = True 
			sWantedDOWs = sWantedDOWs & "1"
		End If 
		If oRs("monday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			If sWantedDOWs <> "" Then 
				sWantedDOWs = sWantedDOWs & ","
			End If 
			sWantedDOWs = sWantedDOWs & "2"
			sDayCell = sDayCell & "Monday"
		End If 
		If oRs("tuesday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			If sWantedDOWs <> "" Then 
				sWantedDOWs = sWantedDOWs & ","
			End If 
			sWantedDOWs = sWantedDOWs & "3"
			sDayCell = sDayCell & "Tuesday"
		End If 
		If oRs("wednesday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			If sWantedDOWs <> "" Then 
				sWantedDOWs = sWantedDOWs & ","
			End If 
			sWantedDOWs = sWantedDOWs & "4"
			sDayCell = sDayCell & "Wednesday"
		End If 
		If oRs("thursday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			If sWantedDOWs <> "" Then 
				sWantedDOWs = sWantedDOWs & ","
			End If 
			sWantedDOWs = sWantedDOWs & "5"
			sDayCell = sDayCell & "Thursday"
		End If 
		If oRs("friday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			If sWantedDOWs <> "" Then 
				sWantedDOWs = sWantedDOWs & ","
			End If 
			sWantedDOWs = sWantedDOWs & "6"
			sDayCell = sDayCell & "Friday"
		End If 
		If oRs("saturday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			If sWantedDOWs <> "" Then 
				sWantedDOWs = sWantedDOWs & ","
			End If 
			sWantedDOWs = sWantedDOWs & "7"
			sDayCell = sDayCell & "Saturday"
		End If 
		response.write "<span class=""daysdisplay"">" &sDayCell & "</span>"
		response.write vbcrlf & "</p>"

		' we need to add a space to the front of the time and before the AM or PM
		sStartTime = " " & Replace(Replace(oRs("starttime"),"AM"," AM"),"PM"," PM")
		sEndTime = " " & Replace(Replace(oRs("endtime"),"AM"," AM"),"PM"," PM")

		If bIsNewRTI Then 
			' Get the days wanted in an array
			iTotalDays = SetWeeklyDates( aWantedDates, sStartDate, sEndDate, "selectedperiod", 0, sStartTime, sEndTime, sWantedDOWs )
			' Save those dates into the temp date table
			SaveClassWantedDates iReservationTempId, aWantedDates, 0, oRs("timedayid")

'			For x = 0 To UBound(aWantedDates, 2) 
'				response.write aWantedDates(0,x) & " &mdash; " & aWantedDates(1,x) & "<br /><br />"
'			Next 
		End If 

		response.write vbcrlf & "<table class=""reservationtempdates"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
		response.write vbcrlf & "<tr><th class=""firstcell"">Include</th><th>Date</th><th>Start Time</th><th>End Time</th><th class=""lastcell"">Available</th></tr>"

		' Display to the page
		ShowRentalAvailabilityDetails iRentalId, iReservationTempId, "selectedperiod", iTotalRows, oRs("timedayid")

		response.write vbcrlf & "</table>"
		oRs.MoveNext 
	Loop

	response.write vbcrlf & "<p>"
	response.write vbcrlf & "<input type=""button"" class=""button"" id=""checkbutton"" name=""checkbutton"" value=""Check Dates"" onclick=""CheckDates()"" />&nbsp; "
	response.write "<input type=""button"" class=""button"" id=""continuebutton"" name=""continuebutton"" value=""Check and Reserve"" onclick=""Validate()"" />"
	response.write vbcrlf & "</p>"
	
	response.write vbcrlf & "<input type=""hidden"" id=""maxrows"" name=""maxrows"" value=""" & iTotalRows & """ />"
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' ShowRentalAvailabilityDetails iRentalId, iReservationTempId, sPeriodTypeSelector, iRowCount, iTimeDayId
'--------------------------------------------------------------------------------------------------
Sub ShowRentalAvailabilityDetails( ByVal iRentalId, ByVal  iReservationTempId, ByVal sPeriodTypeSelector, ByRef iRowCount, ByVal iTimeDayId )
	Dim sSql, oRs, sAmPm, sReservationStartTime, sReservationEndTime, iEndDay, x
	Dim aCheckDates(1,0)

	'For x = 0 To UBound(aWantedDates, 2) 

	' Get the temp dates
	sSql = "SELECT reservationstarttime, reservationendtime, endday FROM egov_rentalreservationdatestemp "
	sSql = sSql & "WHERE reservationtempid = " & iReservationTempId & " AND timedayid = " & iTimeDayId
	sSql = sSql & " ORDER BY reservationstarttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1

		' Build the date time row
		response.write vbcrlf & "<tr class=""dateline"">"
		 
		response.write "<td align=""center"" class=""checkboxcell"">"
		' Checkbox
		response.write "<input type=""checkbox"" id=""includereservationtime" & iRowCount & """ name=""includereservationtime" & iRowCount & """ checked=""checked"" />"
		response.write "<input type=""hidden"" id=""timedayid" & iRowCount & """ name=""timedayid" & iRowCount & """ value=""" & iTimeDayId & """ />"
		response.write "</td>"

		' Start Date
		response.write "<td class=""datecell"">"
		response.write "<input type=""text"" id=""startdate" & iRowCount & """ name=""startdate" & iRowCount & """ value=""" & DateValue(oRs("reservationstarttime")) & """ readonly=""readonly"" size=""10"" maxlength=""10"" onclick=""javascript:void doCalendar('startdate" & iRowCount & "');"" />"
		response.write "&nbsp;<span class=""calendarimg""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('startdate" & iRowCount & "');"" /></span>"
		response.write "</td>"

		sReservationStartTime = oRs("reservationstarttime")
		sReservationEndTime = oRs("reservationendtime")
		iEndDay = oRs("endday")

		response.write "<td align=""center"">" 
		ShowHourPicks "starthour" & iRowCount, GetHourFromDateTime( sReservationStartTime, sAmPm ), ""  ' In rentalscommonfunctions.asp
		response.write ":"
		ShowMinutePicks "startminute" & iRowCount, Minute(sReservationStartTime), ""
		response.write " "
		ShowAmPmPicks "startampm" & iRowCount, sAmPm, ""
		response.write "</td>"

		response.write "<td align=""center"">" 
		ShowHourPicks "endhour" & iRowCount, GetHourFromDateTime( sReservationEndTime, sAmPm ), ""
		response.write ":"
		ShowMinutePicks "endminute" & iRowCount, Minute(sReservationEndTime), ""
		response.write " "
		ShowAmPmPicks "endampm" & iRowCount, sAmPm, ""
		response.write " "
		ShowSameNextDayPick "endday" & iRowCount, iEndDay, ""
		response.write "</td>"

		' Get the availability flag on that date and time
		response.write "<td class=""lastcell"" align=""center"">"
		aCheckDates(0,0) = oRs("reservationstarttime")
		aCheckDates(1,0) = oRs("reservationendtime")
		ShowRentalAvailabilityFlag iRentalId, aCheckDates, sPeriodTypeSelector, True 
		response.write "</td>"
		
		response.write "</tr>"

		response.write vbcrlf & "<tr>"
		response.write "<td class=""firstcell"" colspan=""3"">"
		' Get rental details for that date
		response.write WeekDayName(Weekday(DateValue(oRs("reservationstarttime"))))
		response.write " &ndash; " & GetRentalSeason( iRentalId, DateValue(oRs("reservationstarttime")) )
		response.write GetRentalHours( iRentalId, DateValue(oRs("reservationstarttime")) )
		response.write "</td>"

		response.write "<td class=""lastcell"" colspan=""2"" valign=""top"">Also happening on this date (includes ending buffer)"
		' Get the other reservations, etc for this date
		ShowOtherReservationsForDate iRentalId, DateValue(oRs("reservationstarttime"))
		response.write "</td>"
		response.write "</tr>"

		' The seperator Row
		response.write "<tr><td colspan=""5"" class=""tempseparator"">&nbsp;</td></tr>"
		
		oRs.MoveNext 
	Loop 

	oRs.Close 
	Set oRs = Nothing 
	
End Sub 


'--------------------------------------------------------------------------------------------------
' iReservationTempId = SaveReservationTempInfo( iRentalId, iTimeId )
'--------------------------------------------------------------------------------------------------
Function SaveReservationTempInfo(ByVal iRentalId, ByVal iTimeId )

	Dim sSql, iReservationTempId

	sSql = "DELETE FROM egov_rentalreservationstemp WHERE sessionid = '" & Session.SessionID & "'"
	RunSQLStatement sSql

	sSql = "INSERT INTO egov_rentalreservationstemp ( sessionid, orgid, rentalid, "
	sSql = sSql & " timeid, adminuserid ) VALUES ( '" & Session.SessionID & "', " & session("orgid") & ", " 
	sSql = sSql & iRentalId & ", " & iTimeId & ", " & session("userid") & " )"

	iReservationTempId = RunInsertStatement( sSql )

	SaveReservationTempInfo = iReservationTempId
End Function 


'--------------------------------------------------------------------------------------------------
' SaveClassWantedDates iReservationTempId, aWantedDates, iEndDay
'--------------------------------------------------------------------------------------------------
Sub SaveClassWantedDates( ByVal iReservationTempId, ByRef aWantedDates, ByVal iEndDay, ByVal iTimeDayId )
	Dim sSql, x

	sSql = "DELETE FROM egov_rentalreservationdatestemp WHERE reservationtempid = " & iReservationTempId
	RunSQLStatement sSql

	For x = 0 To UBound(aWantedDates,2)
		sSql = "INSERT INTO egov_rentalreservationdatestemp ( reservationtempid, sessionid, orgid, "
		sSql = sSql & "position, reservationstarttime, reservationendtime, endday, timedayid ) VALUES ( "
		sSql = sSql & iReservationTempId & ", '" & Session.SessionID & "', " & session("orgid") & ", "
		sSql = sSql & x & ", '" & aWantedDates(0,x) & "', '" & aWantedDates(1,x) & "', "
		sSql = sSql & iEndDay & ", " & iTimeDayId & " )"

		RunSQLStatement sSql
	Next 

End Sub 



%>
