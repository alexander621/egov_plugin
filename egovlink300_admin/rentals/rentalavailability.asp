<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalavailability.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals. From here you can create or edit rentals
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iCategoryId, iRentalId, sViewType, sWantedDOWs, s1Checked, s2Checked, s3Checked, s4Checked
Dim s5Checked, s6Checked, s7Checked, sWeeklyDOW, sSelectDate, sStartDate, sEndDate
Dim aWantedDates()

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create simple reservations", sLevel	' In common.asp

If request("rti") = "" Then
	' There is no rti so they are coming from the rental selection page

	iCategoryId = CLng(request("cid"))

	If request("rid") <> "" Then
		iRentalId = CLng(request("rid"))
	Else
		iRentalId = CLng(0) 
	End If 

	If request("viewtype") <> "" Then
		sViewType = request("viewtype")
	Else
		sViewType = "none"
	End If 

	sWantedDOWs = ""
	s1Checked = ""
	s2Checked = ""
	s3Checked = ""
	s4Checked = ""
	s5Checked = ""
	s6Checked = ""
	s7Checked = ""

	If sViewType = "viewselecteddays" Then 
		For Each sWeeklyDOW In request("weeklydow")
			Select Case sWeeklyDOW
				Case "1"
					s1Checked = " checked=""checked"" "
				Case "2"
					s2Checked = " checked=""checked"" "
				Case "3"
					s3Checked = " checked=""checked"" "
				Case "4"
					s4Checked = " checked=""checked"" "
				Case "5"
					s5Checked = " checked=""checked"" "
				Case "6"
					s6Checked = " checked=""checked"" "
				Case "7"
					s7Checked = " checked=""checked"" "
			End Select 
			If sWantedDOWs <> "" Then 
				sWantedDOWs = sWantedDOWs & ","
			End If 
			sWantedDOWs = sWantedDOWs & sWeeklyDOW
		Next 
	End If 

	If request("selectdate") <> "" Then
		sSelectDate = request("selectdate")
	Else
		sSelectDate = ""
	End If 

	If request("startdate") <> "" Then
		sStartDate = request("startdate")
	Else
		sStartDate = ""
	End If 

	If request("enddate") <> "" Then
		sEndDate = request("enddate")
	Else
		sEndDate = ""
	End If 
Else
	' If we have an rti we are returning from the time pick page so get the info we need to show
	iReservationTempId = CLng(request("rti"))

	bHasData = SetPageVariables( iReservationTempId )

	If bHasData Then 
		ClearTempReservation iReservationTempId
	Else
		' Take them somewhere safe, as their data is gone now.
		response.redirect "rentalcategoryselection.asp"
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
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="Javascript">
	<!--
		
		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			var sSelectedDate = '';

			if ($(sField).value != '')
			{
				// The value in the field
				sSelectedDate = $(sField).value;
			}
			else
			{
				if (sField == 'enddate')
				{
					// Show the end date from where the start date is
					sSelectedDate = $("startdate").value;
				}

				if (sSelectedDate == '')
				{
					// This is today's date
					sSelectedDate = new Date();
					var month = sSelectedDate.getMonth() + 1;
					var day = sSelectedDate.getDate();
					var year = sSelectedDate.getFullYear();
					sSelectedDate = month + "/" + day + "/" + year;
				}
			}

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&updatefield=' + sField + '&updateform=frmRentalSearch", "_calendar", "width=350,height=250,toolbar=0,status=0,scrollbars=0,menubar=0,titlebar=0,location=0,dependent=yes,personalbar=no,left=' + w + ',top=' + h + '")');
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

		function viewSingleDate()
		{
			// set the view type
			$("frmviewtype").value = "viewsingledate";

			// make sure the date field is filled in
			if ($("selectdate").value == "")
			{
				displayScreenMsg("Please enter a date, then try viewing again.");
				$("selectdate").focus();
				return false;
			}
			else
			{
				if (! isValidDate($("selectdate").value))
				{
					displayScreenMsg("The date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("selectdate").focus();
					return false;
				}
			}

			// blank out the other dates
			$("startdate").value = "";
			$("enddate").value = "";

			// submit the form
			document.frmRentalSearch.submit();
		}

		function viewSelectedDays()
		{
			var i;
			var hasDOW = false;

			// set the view type
			$("frmviewtype").value = "viewselecteddays";

			// make sure the date fields are filled in
			if ($("startdate").value == "")
			{
				displayScreenMsg("Please enter a start date, then try viewing again.");
				$("startdate").focus();
				return false;
			}
			else
			{
				if (! isValidDate($("startdate").value))
				{
					displayScreenMsg("The start date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("startdate").focus();
					return false;
				}
			}

			if ($("enddate").value == "")
			{
				displayScreenMsg("Please enter an end date, then try viewing again.");
				$("enddate").focus();
				return false;
			}
			else
			{
				if (! isValidDate($("enddate").value))
				{
					displayScreenMsg("The end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("enddate").focus();
					return false;
				}
			}

			// Make use at least one DOW is checked
			for (i=0;i<document.frmRentalSearch.weeklydow.length;i++) 
			{
				if (document.frmRentalSearch.weeklydow[i].checked) 
				{
					hasDOW = true;
				}
			}
			if (hasDOW == false)
			{
				displayScreenMsg('Please select at least one day of the week, then try viewing again.');
				document.frmRentalSearch.weeklydow[0].focus();
				return false;
			}

			// blank out the other date
			$("selectdate").value = "";

			// submit the form
			document.frmRentalSearch.submit();
		}

		function SelectDate( iRentalId, sDate )
		{
			//alert( sDate );
			$("selectedrid").value = iRentalId;
			$("selecteddate").value = sDate;
			document.selectForm.submit();
		}

	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Make Simple Reservations</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->
			<p>
				<span id="screenMsg">&nbsp;</span>
				<input type="button" class="button" value="<< Back" onclick="location.href='rentalofferings.asp?categoryid=<%=iCategoryId%>';" />
			</p>


			<form method="post" name="frmRentalSearch" action="rentalavailability.asp">
				<input type="hidden" id="frmcid" name="cid" value="<%=iCategoryId%>" />
				<input type="hidden" id="frmrid" name="rid" value="<%=iRentalId%>" />
				<input type="hidden" id="frmviewtype" name="viewtype" value="<%=sViewType%>" />

				<p id="selectedname">
			<%	
				If iRentalId > CLng(0) Then
					' Display Rental Info
					response.write "You have selected: " & GetRentalName( iRentalId )
				Else
					' Display Category Info
					response.write "You have selected the category: " & GetCategoryTitle( iCategoryId ) 
				End If 
			%>
				</p>
				<p>
					To find an available date, select from the following options.
				</p>

				<table id="selectchoice" cellpadding="2" cellspacing="0" border="0">
					<tr>
						
						<td class="selecttitle" align="center">
							Pick a specific date.
						</td>
						<td id="orcolumn" align="center" valign="middle">
							OR
						</td>
						<td class="selecttitle" align="center" colspan="2">
							Pick a range of dates and the<br />
							days of the week you wish to view.
						</td>
					</tr>
					<tr>
						<td align="center">
							<input type="text" id="selectdate" name="selectdate" value="<%=sSelectDate%>" size="10" maxlength="10" />
							&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('selectdate');" /></span>
						</td>
						<td>
							&nbsp;
						</td>
						<td align="center" nowrap="nowrap">
							Start Date: 
							<input type="text" id="startdate" name="startdate" value="<%=sStartDate%>" size="10" maxlength="10" />
							&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
						</td>
						<td align="center" nowrap="nowrap">
							End Date: 
							<input type="text" id="enddate" name="enddate" value="<%=sEndDate%>" size="10" maxlength="10" />
							&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
						</td>
					</tr>
					<tr>
						<td>&nbsp;</td>
						<td>&nbsp;</td>
						<td colspan="2" align="center" nowrap="nowrap">
							<input type="checkbox" name="weeklydow" value="1" <%=s1Checked%> />Su
							<input type="checkbox" name="weeklydow" value="2" <%=s2Checked%> />Mo
							<input type="checkbox" name="weeklydow" value="3" <%=s3Checked%> />Tu
							<input type="checkbox" name="weeklydow" value="4" <%=s4Checked%> />We
							<input type="checkbox" name="weeklydow" value="5" <%=s5Checked%> />Th
							<input type="checkbox" name="weeklydow" value="6" <%=s6Checked%> />Fr
							<input type="checkbox" name="weeklydow" value="7" <%=s7Checked%> />Sa
						</td>
					</tr>
					<tr id="selectbutton">
						<td align="center">
							<input type="button" class="button" name="viewsingledate" value="View Single Date" onclick="viewSingleDate();" />
						</td>
						<td>&nbsp;</td>
						<td align="center" colspan="2">
							<input type="button" class="button" name="viewselecteddays" value="View Selected Days" onclick="viewSelectedDays();" />
						</td>
					</tr>

				</table>
			</form>

<%
			If sViewType <> "none" Then
				' Get the days wanted in an array
				If sViewType = "viewsingledate" Then 
					' This is a select date
					ReDim aWantedDates(1,0)
					aWantedDates(0,0) = sSelectDate
					aWantedDates(1,0) = CStr(DateAdd("d", 1, CDate(sSelectDate)))
				Else
					' This is weekly on selected days of the week
					SetWeeklyDates aWantedDates, sStartDate, sEndDate, sWantedDOWs 
				End If 


				' They have pressed a button, so show some results
				ShowRentalAvailability iCategoryId, iRentalId, aWantedDates
			End If 
%>

			<form name="selectForm" method="post" action="rentalcontrol.asp">
				<input type="hidden" id="cid" name="cid" value="<%=iCategoryId%>" />
				<input type="hidden" id="rid" name="rid" value="<%=iRentalId%>" />
				<input type="hidden" id="src" name="src" value="dp" />
				<input type="hidden" id="viewtype" name="viewtype" value="<%=sViewType%>" />
				<input type="hidden" id="selecteddate" name="selecteddate" value="" />
				<input type="hidden" id="selectedrid" name="selectedrid" value="<%=iRentalId%>" />
				<input type="hidden" id="selectdate" name="selectdate" value="<%=sSelectDate%>" />
				<input type="hidden" id="startdate" name="startdate" value="<%=sStartDate%>" />
				<input type="hidden" id="enddate" name="enddate" value="<%=sEndDate%>" />
				<input type="hidden" id="wanteddows" name="wanteddows" value="<%=sWantedDOWs%>" />
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
' void SetWeeklyDates aWantedDates, sStartDate, sEndDate, sWantedDOWs
'--------------------------------------------------------------------------------------------------
Sub SetWeeklyDates( ByRef aWantedDates, ByVal sStartDate, ByVal sEndDate, ByVal sWantedDOWs )
	Dim dTempDate, iTotalDays
	
	' There will always be at least one date, so put that in the array
	iTotalDays = 0
	ReDim aWantedDates(1,0)
	dTempDate = CDate(sStartDate)

	Do While dTempDate <= CDate(sEndDate)
		sWeekDay = CStr(Weekday(dTempDate)) ' get the DOW number 1-7
		If InStr(sWantedDOWs, sWeekDay) > 0 Then 
			' If the dow Is a wanted one Then keep it
			ReDim Preserve aWantedDates(1,iTotalDays)
			' No set time periods 
			aWantedDates(0,iTotalDays) = dTempDate
			aWantedDates(1,iTotalDays) = CStr(DateAdd("d", 1, CDate(dTempDate)))
			iTotalDays = iTotalDays + 1
		End If 
		dTempDate = DateAdd("d",1,dTempDate)
	Loop 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalAvailability iCategoryId, iRentalId, aWantedDates
'--------------------------------------------------------------------------------------------------
Sub ShowRentalAvailability( ByVal iCategoryId, ByVal iRentalId, ByRef aWantedDates )
	Dim sSql, oRs, bOffSeasonFlag, bOkToShow, bHasHours, iCount

	bOkToShow = True
	
	'  AND R.publiccanview = 1 AND R.publiccanreserve = 1
	sSql = "SELECT R.rentalid, rentalname, locationname, ISNULL(width,'') AS width, ISNULL(length,'') AS length, "
	sSql = sSql & " ISNULL(capacity,'') AS capacity, ISNULL(shortdescription,'') AS shortdescription, nocosttorent, "
	sSql = sSql & " ISNULL(iconimageurl,'') AS iconimageurl, publiccanreserve  "
	sSql = sSql & " FROM egov_rentals_list R, egov_rentals_to_categories C "
	sSql = sSql & " WHERE R.rentalid = C.rentalid "
	sSql = sSql & " AND C.recreationcategoryid = " & iCategoryId
	If CLng(iRentalId) > CLng(0) Then
		sSql = sSql & " AND R.rentalid = " & iRentalId
	End If 
	sSql = sSql & " ORDER BY locationname, rentalname"
	'response.write sSql & "<br /><br />"


	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write "<table class=""availablerentals"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
			response.write "<tr><td class=""spacerrow"">&nbsp;</td><td class=""spacerrow"">&nbsp;</td>"   '<td class=""selecttime"">&nbsp;</td>
			response.write "</tr>"
			response.write "<tr>"
'			response.write "<td valign=""top"" class=""iconcell"">" 
'			' Show the name
'			If oRs("iconimageurl") <> "" Then 
'				response.write "<img src=""" & oRs("iconimageurl") & """ alt=""" & oRs("rentalname") & """ title=""" & oRs("rentalname") & """ class=""availabilityimg"" />"
'			Else
'				response.write "&nbsp;"
'			End If 
'			response.write "</td>"

			response.write "<td colspan=""2"" valign=""top"" align=""left"" class=""desccolumn"">"
			response.write "<p><span class=""availableschedulerentalname"">"
			If oRs("locationname")  <> "" Then 
				response.write oRs("locationname") & " &ndash; " 
			End If 
			response.write oRs("rentalname")
			response.write "</span></p>"

			If oRs("shortdescription") <> "" Then 
				response.write "<p>" & oRs("shortdescription") & "</p>"
			End If 
			
			If oRs("width") <> "" Or oRs("capacity") <> "" Then 
				response.write vbcrlf & "<p>"
				If oRs("width") <> "" Then 
					response.write "<strong>Dimensions: </strong>" & oRs("width") & " x " & oRs("length") & "<br />"
				End If 
				If oRs("capacity") <> "" Then 
					response.write "<strong>Capacity: </strong>" & oRs("capacity") & "<br />"
				End If 
				response.write vbcrlf & "</p>"
			End If 

			response.write "</td>"
			response.write "</tr>"

			response.write "<tr><td class=""spacerrow"">&nbsp;</td><td colspan=""2"" class=""spacerrow"">&nbsp;</td></tr>"

			iCount = 0
			' Show the dates here
			For x = 0 To UBound(aWantedDates, 2) 
				iCount = iCount + 1
				If iCount Mod 2 <> 0 Then
					sClass = " class=""altrow"" "
				Else
					sClass = ""
				End If 
				response.write "<tr" & sClass & ">"
				response.write "<td align=""left"" class=""datecolumn"" nowrap=""nowrap"">"
				response.write "<span class=""datedisplay"">" & DateValue(CDate(aWantedDates(0,x))) & " &nbsp;&nbsp; " & WeekDayName(Weekday(CDate(aWantedDates(0,x)))) & "</span>"
				response.write "</td>"
				response.write "<td align=""left"" class=""datecolumn"" colspan=""2""><span class=""availableschedulerentalname"">Availability</span></td></tr>"

				response.write "<tr" & sClass & ">"
				response.write "<td>&nbsp;</td>"
				response.write "<td valign=""top"" align=""left"" class=""availabledatecolumn"" nowrap=""nowrap"">"
				bOffSeasonFlag = GetOffSeasonFlag( oRs("rentalid"), CDate(aWantedDates(0,x)) )

				bHasHours = DisplayAvailability( oRs("rentalid"), CDate(aWantedDates(0,x)), bOffSeasonFlag )
				
				If bHasHours Then 
					response.write "<input type=""button"" class=""button selecttime"" value=""Select this Date and Continue"" onclick=""SelectDate( " & oRs("rentalid") & ", '" & aWantedDates(0,x) & "' );"" />"
				End If 
				response.write "</td>"
				response.write "</tr>"
			Next 
			

			response.write "</table>"
			oRs.MoveNext
		Loop
	Else
		response.write "<p>No Rentals were found.</p>"
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean DisplayAvailability iRentalId, dWantedDate, bOffSeasonFlag
'--------------------------------------------------------------------------------------------------
Function DisplayAvailability( ByVal iRentalId, ByVal dWantedDate, ByVal bOffSeasonFlag )
	Dim dAvailableDate, bHasHours, sPhoneNumber

	sPhoneNumber = GetRentalSupervisorPhone( iRentalId )
	If sPhoneNumber = "" Then
		' Use the City Default Phone Number 
		' sDefaultPhone is set in include_top_functions
		sPhoneNumber = FormatPhoneNumber( sDefaultPhone )	' in common.asp
	End If 
	sPhoneNumber = Trim(sPhoneNumber)

	' see if the date is in the past
	If dWantedDate < CDate(DateValue(Date())) Then
		response.write "<p class=""noreservemsg"">Reservations cannot be made for past dates.</p>"
		bHasHours = False
	Else 
		' check if closed on this date
		If RentalIsClosed( iRentalId, bOffSeasonFlag, Weekday(dWantedDate) ) Then
			response.write "<p class=""noreservemsg"">Closed</p>"
			bHasHours = False
		Else 
			' Go find the available times here
			bHasHours = ShowAvailability( iRentalId, bOffSeasonFlag, Weekday(dWantedDate), dWantedDate )
		End If 
	End If 

	DisplayAvailability = bHasHours

End Function 


'--------------------------------------------------------------------------------------------------
' boolean ShowAvailability( iRentalid, bOffSeasonFlag, iWeekday, dStartDate )
'--------------------------------------------------------------------------------------------------
Function ShowAvailability( ByVal iRentalid, ByVal bOffSeasonFlag, ByVal iWeekday, ByVal dStartDate )
	Dim sSql, oRs, dOpeningTime, dClosingTime, dLastStart, iCount, iMinInterval, sDateAddString, bIsAllDay
	Dim iPostBuffer, iAvailableTimeBlock, dLatestAllowed, dReservationStartTime

	GetOpeningAndClosingTimes iRentalid, bOffSeasonFlag, iWeekday, dStartDate, dOpeningTime, dClosingTime

	' Get Minimal Time interval Info for this day
	bHasMinimum = GetMinimalTimeInfo( iRentalid, bOffSeasonFlag, iWeekday, iMinInterval, sDateAddString, bIsAllDay )

	If Not bIsAllDay Then 
		' convert this interval into minutes if needed
		If sDateAddString = "h" Then
			iMinInterval = CLng(iMinInterval) * 60
		End If 

		' get the post buffer for this day, if any
		'GetPostBufferTime iRentalid, bOffSeasonFlag, iWeekday, iPostBuffer, sDateAddString

		'If sDateAddString = "h" Then
			' convert to minutes
		'	iPostBuffer = CLng(iPostBuffer) * 60
		'End If 

		' This is the admin side so we do not add the buffer in automatically
		iPostBuffer = clng(0)

		' add the post buffer to the minimum allowed time
		'iMinInterval = CLng(iMinInterval) + CLng(iPostBuffer)
	End If 

	sSql = "SELECT reservationstarttime, reservationendtime, billingendtime " 
	sSql = sSql & " FROM egov_rentalreservationdates WHERE rentalid = " & iRentalid
	sSql = sSql & " AND statusid IN (SELECT reservationstatusid FROM egov_rentalreservationstatuses WHERE iscancelled = 0) "
	sSql = sSql & " AND reservationstarttime > '" & DateValue(dStartDate) & " 0:00 AM' "
	sSql = sSql & " AND reservationstarttime < '" & DateValue(DateAdd("d", 1, dStartDate)) & " 0:00 AM' "
	sSql = sSql & " ORDER BY reservationstarttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	dLastStart = dOpeningTime
	iCount = clng(0)
	Do While Not oRs.EOF
		'If clng(iPostBuffer) > clng(0) Then 
		'	dReservationStartTime = DateAdd("n", -(iPostBuffer), CDate(oRs("reservationstarttime")) )
		'Else
			dReservationStartTime = CDate(oRs("reservationstarttime"))
		'End If 

		'response.write "dReservationStartTime = " & dReservationStartTime & "<br />"
		If dLastStart < CDate(oRs("reservationstarttime")) Then 
			If Not bIsAllDay Then 
				'response.write "dLastStart = " & dLastStart & "<br />"
				iAvailableTimeBlock = DateDiff("n", dLastStart, dReservationStartTime)
				'response.write "iAvailableTimeBlock = " & iAvailableTimeBlock & "<br />"
				'response.write "iMinInterval = " & iMinInterval & "<br />"
				' if the available time >= minimum allowed time then output the string
				'If (bHasMinimum = False) Or (bHasMinimum = True And iAvailableTimeBlock >= iMinInterval) Then 
					iCount = iCount + 1
					If iCount > clng(1) Then
						response.write "<br />"
					End If 
					response.write FormatTimeString( dLastStart ) & " to " & FormatTimeString( dReservationStartTime ) '& " - " & iAvailableTimeBlock & " Min"
				'End If 
			End If 
		End If 
		dLastStart = CDate(oRs("billingendtime"))
		oRs.MoveNext
	Loop

	If dLastStart < dClosingTime Then 
		iAvailableTimeBlock = DateDiff("n", dLastStart, dClosingTime)
		'If (bHasMinimum = False) Or (bHasMinimum = True And iAvailableTimeBlock >= iMinInterval) Then 
			' check that the last start time is before the latest reservation start time
			dLatestAllowed = GetLatestReservationTime( iRentalid, bOffSeasonFlag, iWeekday, dStartDate )
'			response.write "<br />" & dLatestAllowed & "<br />"
'			response.write dClosingTime & "<br />"
'			response.write DateDiff("n", dLastStart, dLatestAllowed) & "<br />"
			If DateDiff("n", dLastStart, dLatestAllowed) >= 0 Then 
				iCount = iCount + 1
				If iCount > clng(1) Then
					response.write "<br />"
				End If 
				response.write FormatTimeString( dLastStart ) & " to " & FormatTimeString( dClosingTime ) '& " - " & iAvailableTimeBlock & " Min"
			End If 
		'End If 
	End If 

	If clng(iCount) = clng(0) Then 
		response.write "Unavailable"
		ShowAvailability = False 
	Else 
		ShowAvailability = True 
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' boolean SetPageVariables( iReservationTempId )
'--------------------------------------------------------------------------------------------------
Function SetPageVariables( ByVal iReservationTempId )
	Dim sSql, oRs

	sSql = "SELECT cid, rid, viewtype, rentalid, selectdate, startdate, enddate, weeklydays "
	sSql = sSql & " FROM egov_rentalreservationstemp "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	sSql = sSql & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iCategoryId = CLng(oRs("cid"))
		iRentalId = CLng(oRs("rid"))
		sViewType = oRs("viewtype")
		sWantedDOWs = oRs("weeklydays")
		If sWantedDOWs <> "" Then 
			If sViewType = "viewselecteddays" Then
				If InStr(sWantedDOWs, "1") > 0 Then
					s1Checked = " checked=""checked"" "
				End If 
				If InStr(sWantedDOWs, "2") > 0 Then
					s2Checked = " checked=""checked"" "
				End If
				If InStr(sWantedDOWs, "3") > 0 Then
					s3Checked = " checked=""checked"" "
				End If
				If InStr(sWantedDOWs, "4") > 0 Then
					s4Checked = " checked=""checked"" "
				End If
				If InStr(sWantedDOWs, "5") > 0 Then
					s5Checked = " checked=""checked"" "
				End If
				If InStr(sWantedDOWs, "6") > 0 Then
					s6Checked = " checked=""checked"" "
				End If
				If InStr(sWantedDOWs, "7") > 0 Then
					s7Checked = " checked=""checked"" "
				End If
			End If
		Else
			s1Checked = ""
			s2Checked = ""
			s3Checked = ""
			s4Checked = ""
			s5Checked = ""
			s6Checked = ""
			s7Checked = ""
		End If 
		sSelectDate = oRs("selectdate")
		sStartDate = oRs("startdate")
		sEndDate = oRs("enddate")
		SetPageVariables = True 
	Else
		SetPageVariables = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function  


%>
