<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalsearch.asp
' AUTHOR: Steve Loar
' CREATED: 09/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Search page for rental availability.
'
' MODIFICATION HISTORY
' 1.0   09/28/2009	Steve Loar - INITIAL VERSION
' 1.1	03/21/2011	Steve Loar - Adding ability to add dates to existing reservations
' 1.2	03/24/2011	Steve Loar - hide deactivated rentals
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, iSearchItem, iRecreationCategoryId, iLocationId
Dim sRentalName, sStartDate, sEndDate, iCitizenUserid, iPeriodTypeId, iStartHour, iEndHour
Dim iStartMinute, sStartAmPm, iEndMinute, sEndAmPm, sOChecked, sDChecked, sWChecked
Dim sMChecked, s1Checked, s2Checked, s3Checked, s4Checked, s5Checked, s6Checked, s7Checked
Dim sWeeklyDOW, iMonthlyPeriod, iMonthlyDOW, iOrderBy, bOkToSearch, sLoadMsg, dEndTime
Dim dStartTime, iEndDay, sStartTime, sEndTime, sWantedDOWs, iTotalDays, sPeriodType
Dim iReservationTempId, iReservationId, sSubTitle
Dim aWantedDates()

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "make reservations", sLevel	' In common.asp

sWantedDOWs = ""
iTotalDays = CLng(0) 
iReservationTempId = 0

If request("rid") <> "" Then
	iReservationId = CLng(request("rid"))
	If iReservationId > CLng(0) Then 
		sSubTitle = " &ndash; Add Dates to an Existing Reservation"
		' confirm that rid is in their org and redirect if not
		If Not OrgHasReservation( iReservationId ) Then
			response.redirect request.ServerVariables("HTTP_REFERER")
		End If 
	Else
		sSubTitle = " "
	End If 
Else
	iReservationId = CLng(0)
	sSubTitle = " "
End If 

If request("recreationcategoryid") <> "" Then
	iRecreationCategoryId = CLng(request("recreationcategoryid"))
Else
	iRecreationCategoryId = 0
End If 

If request("locationid") <> "" Then
	iLocationId = CLng(request("locationid"))
Else
	iLocationId = 0
End If 

If request("rentalname") <> "" Then
	sRentalName = request("rentalname")
Else
	sRentalName = ""
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

If request("citizenuserid") <> "" Then 
	iCitizenUserid = CLng(request("citizenuserid"))
Else
	iCitizenUserid = CLng(0)
End If 

If request("periodtypeid") <> "" Then 
	iPeriodTypeId = CLng(request("periodtypeid"))
Else
	iPeriodTypeId = CLng(0)
End If 

If request("starthour") <> "" Then
	iStartHour = request("starthour")
Else
	iStartHour = "1"
End If 

If request("startminute") <> "" Then
	iStartMinute = request("startminute")
Else
	iStartMinute = "00"
End If 

If request("startampm") <> "" Then
	sStartAmPm = request("startampm")
Else
	sStartAmPm = "AM"
End If 

If request("endhour") <> "" Then
	iEndHour = request("endhour")
Else
	iEndHour = "1"
End If 

If request("endminute") <> "" Then
	iEndMinute = request("endminute")
Else
	iEndMinute = "00"
End If 

If request("endampm") <> "" Then
	sEndAmPm = request("endampm")
Else
	sEndAmPm = "AM"
End If 

If request("endday") <> "" Then
	iEndDay = clng(request("endday"))
Else
	iEndDay = clng(0)
End If 

sOChecked = ""
sDChecked = ""
sWChecked = ""
sMChecked = ""
If request("occurs") <> "" Then
	sOccurs = request("occurs")
	Select Case request("occurs")
		Case "o"
			sOChecked = " checked=""checked"" "
		Case "d"
			sDChecked = " checked=""checked"" "
		Case "w"
			sWChecked = " checked=""checked"" "
		Case "m"
			sMChecked = " checked=""checked"" "
	End Select 
Else
	sOccurs =  "o"
	sOChecked = " checked=""checked"" "
End If 

s1Checked = ""
s2Checked = ""
s3Checked = ""
s4Checked = ""
s5Checked = ""
s6Checked = ""
s7Checked = ""
If sWChecked <> "" Then 
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

If request("monthlyperiodid") <> "" Then
	iMonthlyPeriod = CLng(request("monthlyperiodid"))
Else
	iMonthlyPeriod = CLng(0)
End If 

If request("monthlydow") <> "" Then
	iMonthlyDOW = clng(request("monthlydow"))
Else
	iMonthlyDOW = clng(0)
End If 

If request("orderby") <> "" Then
	iOrderBy = clng(request("orderby"))
Else
	iOrderBy = clng(1)
End If 

'response.write request.servervariables("REQUEST_METHOD") & "<br />"
If request.servervariables("REQUEST_METHOD") = "POST" Then 
	'response.write "post<br />"
	' Handle the post back from the dateselection page
	If request("rti") <> "" AND request("weeklydays") <> "" Then
		sWantedDOWs = request("weeklydays")
		If InStr(request("weeklydays"),"1") > 0 Then
			s1Checked = " checked=""checked"" "
		End If 
		If InStr(request("weeklydays"),"2") > 0 Then
			s2Checked = " checked=""checked"" "
		End If 
		If InStr(request("weeklydays"),"3") > 0 Then
			s3Checked = " checked=""checked"" "
		End If 
		If InStr(request("weeklydays"),"4") > 0 Then
			s4Checked = " checked=""checked"" "
		End If 
		If InStr(request("weeklydays"),"5") > 0 Then
			s5Checked = " checked=""checked"" "
		End If 
		If InStr(request("weeklydays"),"6") > 0 Then
			s6Checked = " checked=""checked"" "
		End If 
		If InStr(request("weeklydays"),"7") > 0 Then
			s7Checked = " checked=""checked"" "
		End If 
	End If 

	' validate the picks before allowing a search to happen

	sPeriodType = GetSelectedPeriodType( iPeriodTypeId )

	'dEndDate = CDate(sEndDate & " " & iEndHour & ":" & iEndMinute & " " & sEndAmPm)
	'dStartDate = CDate(sStartDate & " " & iStartHour & ":" & iStartMinute & " " & sStartAmPm)

	dStartTime = CDate(sStartDate & " " & iStartHour & ":" & iStartMinute & " " & sStartAmPm)
	sStartTime = " " & iStartHour & ":" & iStartMinute & " " & sStartAmPm
	If iEndDay = clng(0) Then 
		dEndTime = CDate(sStartDate & " " & iEndHour & ":" & iEndMinute & " " & sEndAmPm)
	Else
		dEndTime = CDate(CStr(DateAdd("d", 1, CDate(sStartDate))) & " " & iEndHour & ":" & iEndMinute & " " & sEndAmPm)
	End If 
	' check that the selected end time is at least the mininum rental period. This is on an Org level.
	'CheckMinimumRentalInterval dStartTime, dEndTime, iEndHour, iEndMinute, sEndAmPm 
	CheckOrgRentalRoundUp dStartTime, dEndTime, iEndHour, iEndMinute, sEndAmPm

	sEndTime = " " & iEndHour & ":" & iEndMinute & " " & sEndAmPm

	' Make sure the end date is not earlier than the start date
	If CDate(sEndDate) < CDate(sStartDate) Then 
		bOkToSearch = False  
		sLoadMsg = "displayScreenMsg('The end date cannot be earlier than the start date. Please change this, then try your search again.');"
		sLoadMsg = sLoadMsg & vbcrlf & "$(""enddate"").focus();"
	' Make sure the end time is not earlier than the start time
	ElseIf IsSelectedTimePeriod( iPeriodTypeId ) And dEndTime <= dStartTime Then 
		bOkToSearch = False  
		sLoadMsg = "displayScreenMsg('The end time must be later than the start time. Please change this, then try your search again.');"
		sLoadMsg = sLoadMsg & vbcrlf & "$(""endhour"").focus();"
	Else 
		bOkToSearch = True 
		'sLoadMsg = "displayScreenMsg('Is OK.');"
	End If 
'Else 
	'sLoadMsg = "displayScreenMsg('Not a Post.');"
End If 

			If bOkToSearch Then

				' Get the days wanted in an array
				iTotalDays = GetWantedDates( aWantedDates, sStartDate, sEndDate, sPeriodType, sStartTime, sEndTime, iEndDay, sOccurs, sWantedDOWs, iMonthlyPeriod, iMonthlyDOW )

				' Save the Reservation Details for the date selection page
				iReservationTempId = SaveReservationTempInfo( sStartDate, sEndDate, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm, iEndDay, sOccurs, sWantedDOWs, iMonthlyPeriod, iMonthlyDOW, sSearchName, sRentalName, iLocationId, iRecreationCategoryId, iPeriodTypeId, iOrderBy, iReservationId )
				
				' Save the Wanted Dates for the date selection page
				SaveWantedDates iReservationTempId, aWantedDates, iEndDay

'				For x = 0 To UBound(aWantedDates, 2) 
'					response.write aWantedDates(0,x) & " &mdash; " & aWantedDates(1,x) & "<br /><br />"
'				Next 

			End If 
%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="JavaScript" src="../scripts/setfocus.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="Javascript">
	<!--

		//The two functions handle the row highlight and un-highlight for result lists when the mouse cursor moves over and off a record
		function mouseOverSearchRow( oRow, sNext ) {
			var oNextRow;
			oRow.style.backgroundColor='#93bee1';
			oRow.style.cursor='pointer';

			if (sNext == 'next')
			{
				oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
				if (oNextRow) {
				oNextRow.style.backgroundColor='#93bee1';
				}
			}
			else
			{
				oNextRow = document.getElementById(eval(parseInt(oRow.id) - 1));
				if (oNextRow) {
					oNextRow.style.backgroundColor='#93bee1';
				}
			}
		}

		function mouseOutSearchRow( oRow, sNext ) {	
			oRow.style.backgroundColor='';
			oRow.style.cursor='';

			if (sNext == 'next')
			{
				oNextRow = document.getElementById(eval(parseInt(oRow.id) + 1));
				if (oNextRow) 
				{
					oNextRow.style.backgroundColor='';
				}
			}
			else
			{
				oNextRow = document.getElementById(eval(parseInt(oRow.id) - 1));
				if (oNextRow) 
				{
					oNextRow.style.backgroundColor='';
				}
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

		function doCalendar( sField ) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			var sSelectedDate = $(sField).value;

			// Set the end date to the start date
			if (sField == "enddate")
			{
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

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmRentalSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function Search()
		{
			var i;
			var hasDOW = false; 

			// check for a start date
			if ($("startdate").value == "")
			{
				displayScreenMsg('Please enter a start date, then try your search again.');
				$("startdate").focus();
				return false;
			}
			else
			{
				if (! isValidDate($("startdate").value))
				{
					alert("The start date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("startdate").focus();
					return false;
				}
			}

			// Handle the just once radio pick
			for (i=0;i<document.frmRentalSearch.occurs.length;i++) {
				if (document.frmRentalSearch.occurs[i].checked) {
					var selected_occurs = document.frmRentalSearch.occurs[i].value;
				}
			}

			if (selected_occurs == 'o')
			{
				// set the end date to start date for just this once choice
				$("enddate").value = $("startdate").value;
			}

			// check for an end date
			if ($("enddate").value == "")
			{
				displayScreenMsg('Please enter an end date, then try your search again.');
				$("enddate").focus();
				return false;
			}
			else
			{
				if (! isValidDate($("enddate").value))
				{
					alert("The end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("enddate").focus();
					return false;
				}
			}

			if (selected_occurs == 'w')
			{
				for (i=0;i<document.frmRentalSearch.weeklydow.length;i++) 
				{
					if (document.frmRentalSearch.weeklydow[i].checked) 
					{
						hasDOW = true;
					}
				}
				if (hasDOW == false)
				{
					displayScreenMsg('Please select at least one day of the week, then try your search again.');
					document.frmRentalSearch.weeklydow[0].focus();
					return false;
				}
			}

			document.frmRentalSearch.submit();
		}

		function loader()
		{
<%
    sLoadMsg = sLoadMsg & "changeSelectedRentalIDs();"
    response.write sLoadMsg
%>

		}

  function changeSelectedRentalIDs() {
     var lcl_total_rows         = document.getElementById('total_rentals').value;
     var lcl_selected_rentalids = '';
     var i = 0;

     for (i=1; i<=lcl_total_rows; i++) {
        if(document.getElementById('rentalid' + i).checked) {
           if(lcl_selected_rentalids == '') {
              lcl_selected_rentalids = document.getElementById('rentalid' + i).value;
           } else {
              lcl_selected_rentalids = lcl_selected_rentalids + ',' + document.getElementById('rentalid' + i).value;
           }
        }
     }

     document.getElementById('selected_rentalids').value = lcl_selected_rentalids;

     if(document.getElementById('continueButton')) {
        if(lcl_selected_rentalids != '') {
           document.getElementById('continueButton').disabled = false;
        } else {
           document.getElementById('continueButton').disabled = true;
        }
     }

  }

  function viewRentalDateSelection() {
     var lcl_url       = '';
     var lcl_rentalids = document.getElementById('selected_rentalids').value;

     lcl_url  = 'rentaldateselection.asp';
     lcl_url += '?selected_rentalids=' + lcl_rentalids;
     lcl_url += '&rti=<%=iReservationTempId%>';

     location.href = lcl_url;
  }

  function selectRentalRowClick(iRowID) {
     var lcl_rowID    = Number(0);
     var lcl_rentalid = Number(0);

     if(iRowID != '' && iRowID != undefined) {
        lcl_rowID = Number(iRowID);
     }

     if(lcl_rowID > 0) {
        document.getElementById('rentalid' + lcl_rowID).click();
     }
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
				<font size="+1"><strong>Rental Search<%=sSubTitle%></strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<span id="screenMsg">&nbsp;</span>

			<!--BEGIN: FILTER SELECTION-->
			<div class="rentalsearchselection">
				<fieldset class="filterselection" id="searchfieldset">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmRentalSearch" method="POST" action="rentalsearch.asp">
      
							<table id="rentalsearch" cellpadding="1" cellspacing="0" border="0">
								<tr>
									<td class="labelcolumn">Looking For:</td>
									<td class="pickcolumn"><% ShowRecreationCategories iRecreationCategoryId %></td>
									<!--<td class="labelcolumn2">Occurring: </td>-->
									<td>Reservation Will Occur:<br />
									<input type="radio" name="occurs" value="o" <%=sOChecked%> /> Just This Once</td>
								</tr>
								<tr>
									<td class="labelcolumn">Located At:</td>
									<td class="pickcolumn"><% ShowRentalLocations iLocationId %></td>
									<!--<td class="labelcolumn2">&nbsp;</td>-->
									<td><input type="radio" name="occurs" <%=sDChecked%> value="d" /> Daily
									</td>
								</tr>
								<tr>
									<td class="labelcolumn">Rental Name Like:</td>
									<td class="pickcolumn"><input type="text" id="rentalname" name="rentalname" value="<%=sRentalName%>" size="40" maxlength="40" /></td>
									<!--<td class="labelcolumn2">&nbsp;</td>-->
									<td><input type="radio" name="occurs" <%=sWChecked%> value="w" /> Weekly On These Days
									</td>
								</tr>
								<tr>
									<td class="labelcolumn">Start Date:</td>
									<td class="pickcolumn"><input type="text" id="startdate" name="startdate" value="<%=sStartDate%>" size="10" maxlength="10" />
										&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
									</td>
									<!--<td class="labelcolumn2">&nbsp;</td>-->
									<td>
										<input type="checkbox" name="weeklydow" value="1" <%=s1Checked%> />Su
										<input type="checkbox" name="weeklydow" value="2" <%=s2Checked%> />Mo
										<input type="checkbox" name="weeklydow" value="3" <%=s3Checked%> />Tu
										<input type="checkbox" name="weeklydow" value="4" <%=s4Checked%> />We
										<input type="checkbox" name="weeklydow" value="5" <%=s5Checked%> />Th
										<input type="checkbox" name="weeklydow" value="6" <%=s6Checked%> />Fr
										<input type="checkbox" name="weeklydow" value="7" <%=s7Checked%> />Sa
									</td>
								</tr>
								<tr>
									<td class="labelcolumn">End Date:</td>
									<td class="pickcolumn"><input type="text" id="enddate" name="enddate" value="<%=sEndDate%>" size="10" maxlength="10" />
										&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
									</td>
									<!--<td class="labelcolumn2">&nbsp;</td>-->
									<td><input type="radio" name="occurs" <%=sMChecked%> value="m" /> Monthly On The 
										<% ShowMonthlyPeriodPicks iMonthlyPeriod %>
									</td>
								</tr>
								<tr>
									<td class="labelcolumn">Time Period: </td>
									<td class="pickcolumn"><% ShowRentalPeriods iPeriodTypeId %></td>
									<!--<td class="labelcolumn2">&nbsp;</td>-->
									<td>
										<% ShowDOWPicks "monthlydow", iMonthlyDOW %>
									</td>
								</tr>
								<tr>
									<td class="labelcolumn">Start Time: </td>
									<td class="pickcolumn">
										<% ShowHourPicks "starthour",  iStartHour, "" %>:
										<% ShowMinutePicks "startminute", iStartMinute, "" %>&nbsp;
										<% ShowAmPmPicks "startampm", sStartAmPm, "" %>
									</td>
									<!--<td class="labelcolumn2">&nbsp;</td>-->
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td class="labelcolumn">End Time: </td>
									<td class="pickcolumn"><% ShowHourPicks "endhour",  iEndHour, "" %>:
										<% ShowMinutePicks "endminute", iEndMinute, "" %>&nbsp;
										<% ShowAmPmPicks "endampm", sEndAmPm, "" %>&nbsp;
										<% ShowSameNextDayPick "endday", iEndDay, "" %>
									</td>
									<!--<td class="labelcolumn2">Order By:</td>-->
									<td>Order Results By:<br /><% ShowOrderByPicks iOrderBy %></td>
								</tr>
								<tr>
									<td class="labelcolumn">&nbsp;</td>
									<td class="pickcolumn">
										&nbsp;
									</td>
									<!--<td class="labelcolumn2">&nbsp;</td>-->
									<td>&nbsp;</td>
								</tr>
								<tr>
			    					<td colspan="3"><input class="button" type="button" name="searchbutton" value="Search" onclick="Search();" /></td>
  								</tr>
								</table>
								<input type="hidden" id="rid" name="rid" value="<%=iReservationId%>" />
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->

<%				'Pull the list here
			if bOkToSearch then

				'Get the days wanted in an array
     'iTotalDays = GetWantedDates( aWantedDates, sStartDate, sEndDate, sPeriodType, sStartTime, sEndTime, iEndDay, sOccurs, sWantedDOWs, iMonthlyPeriod, iMonthlyDOW )

				'Save the Reservation Details for the date selection page
     'iReservationTempId = SaveReservationTempInfo( sStartDate, sEndDate, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm, iEndDay, sOccurs, sWantedDOWs, iMonthlyPeriod, iMonthlyDOW, sSearchName, sRentalName, iLocationId, iRecreationCategoryId, iPeriodTypeId, iOrderBy, iReservationId )
				
				'Save the Wanted Dates for the date selection page
     'SaveWantedDates iReservationTempId, aWantedDates, iEndDay

     'For x = 0 To UBound(aWantedDates, 2) 
        'response.write aWantedDates(0,x) & " &mdash; " & aWantedDates(1,x) & "<br /><br />"
     'Next 

				'Look for rentals that match the search criteria
				 SearchForRentals iRecreationCategoryId, _
                      iLocationId, _
                      sRentalName, _
                      iOrderBy, _
                      aWantedDates, _
                      sPeriodType, _
                      iReservationTempId
   else
      response.write "<input type=""hidden"" name=""selected_rentalids"" id=""selected_rentalids"" value="""" size=""10"" maxlength=""500"" />" & vbcrlf
      response.write "<input type=""hidden"" name=""total_rentals"" id=""total_rentals"" value=""0"" size=""3"" maxlength=""50"" />" & vbcrlf
   end if
%>			

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
' void SearchForRentals iRecreationCategoryId, iLocationId, sRentalName, aWantedDates
'--------------------------------------------------------------------------------------------------
Sub SearchForRentals( ByVal iRecreationCategoryId, ByVal iLocationId, ByVal sRentalName, ByVal iOrderBy, ByRef aWantedDates, ByVal sPeriodType, ByVal iReservationTempId )
	Dim sSql, oRs, iRowCount, iRealRowCount

	iRowCount = 0
	iRealRowCount = 0

	sSQL = "SELECT "
 sSQL = sSQL & " rentalid, "
 sSQL = sSQL & " rentalname, "
 sSQL = sSQL & " locationname, "
 sSQL = sSQL & " ISNULL(width,'') AS width, "
 sSQL = sSQL & " ISNULL(length,'') AS length, "
	sSQL = sSQL & " ISNULL(capacity,'') AS capacity, "
 sSQL = sSQL & " ISNULL(shortdescription,'') AS shortdescription, "
 sSQL = sSQL & " nocosttorent "
	sSQL = sSQL & " FROM egov_rentals_list "
 sSQL = sSQL & " WHERE isdeactivated = 0 "
 sSQL = sSQL & " AND orgid = " & session("orgid")

	If CLng(iLocationId) > CLng(0) Then 
  		sSQL = sSQL & " AND locationid = " & iLocationId
	End If 

	If sRentalName <> "" Then
  		sSQL = sSQL & " AND rentalname LIKE '%" & dbsafe(sRentalName) & "%'"
	End If 

 sSQL = sSQL & " AND rentalid IN (SELECT rentalid "
 sSQL = sSQL &                  " FROM egov_rentals_to_categories "
 sSQL = sSQL &                  " WHERE recreationcategoryid = " & iRecreationCategoryId & ") "

	If clng(iOrderBy) = clng(1) Then
  		sSQL = sSQL & " ORDER BY locationname, rentalname"
	Else
  		sSQL = sSQL & " ORDER BY rentalname, locationname"
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

 response.write "<span style=""color:#ff0000"">" & vbcrlf
 response.write "  1. Select rental(s) below.<br />" & vbcrlf
 response.write "  2. To check date availability click... <input type=""button"" name=""continueButton"" id=""continueButton"" value=""Continue"" class=""button"" onclick=""viewRentalDateSelection();"" />" & vbcrlf
 response.write "  <input type=""hidden"" name=""selected_rentalids"" id=""selected_rentalids"" value="""" size=""10"" maxlength=""500"" />" & vbcrlf
 response.write "</span>" & vbcrlf
	response.write "<div id=""rentalresultsshadow"" class=""shadow"">"
	response.write "<table id=""rentalresults"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
	response.write "<tr><th colspan=""2"">Rental</th><th>Location</th><th>Dimensions</th><th>Capacity</th><th>Available</th></tr>"

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			iRowCount        = iRowCount + 1
			iRealRowCount    = iRealRowCount + 1
   lcl_tr1_class    = ""
   lcl_tr2_class    = ""
   lcl_td_onclick   = ""
   lcl_dimensions   = "&nbsp;"
   lcl_nocosttorent = ""

			if iRowCount mod 2 = 0 then
  				lcl_tr1_class = " class=""altrow"""
			end if

   'lcl_td_onclick = "location.href='"
   'lcl_td_onclick = lcl_td_onclick & "rentaldateselection.asp"
   'lcl_td_onclick = lcl_td_onclick & "?rentalid=" & oRs("rentalid")
   'lcl_td_onclick = lcl_td_onclick & "&rti=" & iReservationTempId
   'lcl_td_onclick = lcl_td_onclick & "';"

   lcl_td_onclick = "selectRentalRowClick('" & iRowCount & "');"

  'Set up Dimensions
			if oRs("width") <> "" AND oRs("length") <> "" then
      lcl_dimensions = oRs("width") & " by " & oRs("length")
			end if

  'BEGIN: Row 1 ---------------------------------------------------------------
			response.write "  <tr id=""" & iRealRowCount & """" & lcl_tr1_class & " align=""center"" onMouseOver=""mouseOverSearchRow(this, 'next');"" onMouseOut=""mouseOutSearchRow(this, 'next');"">" & vbcrlf
			response.write "      <td align=""center"" rowspan=""2"">" & vbcrlf
   response.write "          &nbsp;<input type=""checkbox"" name=""rentalid" & iRowCount & """ id=""rentalid" & iRowCount & """ value=""" & oRs("rentalid") & """ onclick=""changeSelectedRentalIDs()"" />&nbsp;" & vbcrlf
   response.write "      </td>" & vbcrlf
			response.write "      <td title=""click to view"" onclick=""" & lcl_td_onclick & """ nowrap=""nowrap"" class=""firstcol"" align=""left""><strong>" & oRs("rentalname") & "</strong></td>" & vbcrlf
			response.write "      <td title=""click to view"" onclick=""" & lcl_td_onclick & """ nowrap=""nowrap""><strong>" & oRs("locationname") & "</strong></td>" & vbcrlf
			response.write "      <td title=""click to view"" onclick=""" & lcl_td_onclick & """ nowrap=""nowrap"">" & lcl_dimensions  & "</td>" & vbcrlf
			response.write "      <td title=""click to view"" onclick=""" & lcl_td_onclick & """ nowrap=""nowrap"">" & oRs("capacity") & "</td>" & vbcrlf
			response.write "      <td title=""click to view"" onclick=""" & lcl_td_onclick & """ nowrap=""nowrap"">" & vbcrlf
   response.write "          <strong>" & vbcrlf
                             ShowRentalAvailabilityFlag oRs("rentalid"), aWantedDates, sPeriodType, False 
			response.write "          </strong>" & vbcrlf
   response.write "      </td>" & vbcrlf
			response.write "  </tr>" & vbcrlf
  'END: Row 1 -----------------------------------------------------------------

  'BEGIN: Row 2 ---------------------------------------------------------------
			iRealRowCount = iRealRowCount + 1

			if iRowCount mod 2 = 0 then
  				lcl_tr2_class = " class=""altrow"""
			end if

			if oRs("nocosttorent") then
      lcl_nocosttorent = "<strong>There is no cost to rent this.</strong><br />" & vbcrlf
			end if

			response.write "  <tr id=""" & iRealRowCount & """" & lcl_tr2_class & " onMouseOver=""mouseOverSearchRow(this, 'prior');"" onMouseOut=""mouseOutSearchRow(this, 'prior');"">" & vbcrlf
 		response.write "      <td colspan=""5"" class=""firstcol2ndrow"" align=""left"" title=""click to view"" onclick=""" & lcl_td_onclick & """>" & vbcrlf
   response.write            lcl_nocosttorent
			response.write            oRs("shortdescription") & vbcrlf
			response.write "      </td>" & vbcrlf
			response.write "  </tr>" & vbcrlf
  'END: Row 2 -----------------------------------------------------------------

			oRs.MoveNext
		Loop
	Else
		response.write vbcrlf & "<tr><td colspan=""5"" class=""firstcol"" align=""left""><strong>No Rentals could be that matched your search criteria.</strong></td></tr>" & vbcrlf
	End If 

	response.write "</table>" & vbcrlf
 response.write "</div>" & vbcrlf
 response.write "<input type=""hidden"" name=""total_rentals"" id=""total_rentals"" value=""" & iRowCount & """ size=""3"" maxlength=""50"" /><br />" & vbcrlf
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowOrderByPicks iOrderBy
'--------------------------------------------------------------------------------------------------
Sub ShowOrderByPicks( ByVal iOrderBy )
	
	response.write "<select name='orderby' id='orderby'>"

	response.write vbcrlf & "<option value='1'"
	If iOrderBy = clng(1) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Location, Name</option>"

	response.write vbcrlf & "<option value='2'"
	If iOrderBy = clng(2) Then
		response.write " selected=""selected"" "
	End If 
	response.write ">Name, Location</option>"

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean IsSelectedTimePeriod( iPeriodTypeId )
'--------------------------------------------------------------------------------------------------
Function IsSelectedTimePeriod( ByVal iPeriodTypeId )
	Dim oRs, sSql

	sSql = "SELECT isselectedperiod FROM egov_rentalperiodtypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND periodtypeid = " & iPeriodTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isselectedperiod") Then
			IsSelectedTimePeriod = True 
		Else
			IsSelectedTimePeriod = False 
		End If 
	Else
		IsSelectedTimePeriod = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer SaveReservationTempInfo( sStartDate, sEndDate, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm, iEndDay, sOccurs, sWantedDOWs, iMonthlyPeriod, iMonthlyDOW, iReservationId )
'--------------------------------------------------------------------------------------------------
Function SaveReservationTempInfo( ByVal sStartDate, ByVal sEndDate, ByVal iStartHour, ByVal iStartMinute, ByVal sStartAmPm, _
	ByVal iEndHour, ByVal iEndMinute, ByVal sEndAmPm, ByVal iEndDay, ByVal sOccurs, ByVal sWantedDOWs, ByVal iMonthlyPeriod, _
	ByVal iMonthlyDOW, ByVal sSearchName, ByVal sRentalName, ByVal iLocationId, ByVal iRecreationCategoryId, ByVal iPeriodTypeId, _
	ByVal iOrderBy, ByVal iReservationId )

	Dim sSql, iReservationTempId

	sSql = "DELETE FROM egov_rentalreservationstemp WHERE sessionid = '" & Session.SessionID & "'"
	RunSQLStatement sSql

	sSql = "DELETE FROM egov_rentalreservationdatestemp WHERE sessionid = '" & Session.SessionID & "'"
	RunSQLStatement sSql

	sSql = "INSERT INTO egov_rentalreservationstemp ( sessionid, orgid, requestedstartdate, "
	sSql = sSql & "requestedstarthour, requestedstartminute, requestedstartampm, requestedenddate, requestedendhour, "
	sSql = sSql & "requestedendminute, requestedendampm, requestedendday, occurs, weeklydays, rentalmonthlyperiodid, "
	sSql = sSql & "monthlydow, adminuserid, userlike, rentallike, locationid, recreationcategoryid, periodtypeid, "
	sSql = sSql & "orderby, reservationid ) VALUES ( '" & Session.SessionID & "', " & session("orgid") & ", " 
	sSql = sSql & "'" & sStartDate & "', " & iStartHour & ", " & iStartMinute & ", '" & sStartAmPm & "', '" & sEndDate & "', "
	sSql = sSql & iEndHour & ", " & iEndMinute & ", '" & sEndAmPm & "', " & iEndDay & ", '" & sOccurs & "', '" & sWantedDOWs & "', "
	sSql = sSql & iMonthlyPeriod & ", "& iMonthlyDOW & ", " & session("userid") & ", '" & dbsafe(sSearchName) & "', '"
	sSql = sSql & dbsafe(sRentalName) & "', "  & iLocationId & ", " & iRecreationCategoryId & ", " & iPeriodTypeId & ", "
	sSql = sSql & iOrderBy & ", " & iReservationId & " )"

	iReservationTempId = RunInsertStatement( sSql )

	SaveReservationTempInfo = iReservationTempId

End Function 


'--------------------------------------------------------------------------------------------------
' void SaveWantedDates iReservationTempId, aWantedDates, iEndDay
'--------------------------------------------------------------------------------------------------
Sub SaveWantedDates( ByVal iReservationTempId, ByRef aWantedDates, ByVal iEndDay )
	Dim sSql, x

	For x = 0 To UBound(aWantedDates,2)
		sSql = "INSERT INTO egov_rentalreservationdatestemp ( reservationtempid, sessionid, orgid, "
		sSql = sSql & "position, reservationstarttime, reservationendtime, endday ) VALUES ( "
		sSql = sSql & iReservationTempId & ", '" & Session.SessionID & "', " & session("orgid") & ", "
		sSql = sSql & x & ", '" & aWantedDates(0,x) & "', '" & aWantedDates(1,x) & "', " & iEndDay & " )"

		RunSQLStatement sSql
	Next 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean OrgHasReservation( iReservationId )
'--------------------------------------------------------------------------------------------------
Function OrgHasReservation( ByVal iReservationId )
	Dim oRs, sSql

	sSql = "SELECT COUNT(reservationid) AS hits FROM egov_rentalreservations "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND reservationid = " & iReservationId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			OrgHasReservation = True 
		Else
			OrgHasReservation = False 
		End If 
	Else
		OrgHasReservation = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
sub dtb_debug(p_value)
  sSqli = "INSERT INTO my_table_dtb (notes) VALUES ('" & replace(p_value,"'","''") & "')"
  set rsi = Server.CreateObject("ADODB.Recordset")
 	rsi.Open sSqli, Application("DSN"), 3, 1
end sub
%>
