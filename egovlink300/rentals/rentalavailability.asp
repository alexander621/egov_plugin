<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="rentalcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentalavailability.asp
' AUTHOR: Steve Loar
' CREATED: 01/19/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Shows Availability by rental, then days selected
'
' MODIFICATION HISTORY
' 1.0   01/19/2010	Steve Loar - INITIAL VERSION
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

Dim sTitle, iCategoryId, iRentalId, sSelectDate, sWeeklyDOW, sWantedDOWs, sViewType
Dim s1Checked, s2Checked, s3Checked, s4Checked, s5Checked, s6Checked, s7Checked
Dim sStartDate, sEndDate, iReservationTempId, bHasData, re, matches
Dim aWantedDates()

Set re = New RegExp
re.Pattern = "^\d+$"

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

If request("rti") = "" Then 
	' This is a post from the categories page
	If request("cid") = "" Then
		response.redirect "rentalcategories.asp"
	Else
		If Not IsNumeric(request("cid")) Then
			response.redirect "rentalcategories.asp"
		Else 
			Set matches = re.Execute(request("cid"))
			If matches.Count > 0 Then
				iCategoryId = CLng(request("cid"))
			Else 
				response.redirect "rentalcategories.asp"
			End If 
		End If 
	End If

	If request("rid") <> "" Then
		Set matches = re.Execute(request("rid"))
		If matches.Count > 0 Then
			iRentalId = CLng(request("rid"))
		Else 
			response.redirect "rentalcategories.asp"
		End If 
	Else
		' setting the rental id to 0 allows searching for the entire category
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
	'response.write "sWantedDOWs: " & sWantedDOWs & "<br />"


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
	' This is a return from another rentals page and we can look up what we need for the form

	' Handle the spiders, and hacks
	If Not IsNumeric(request("rti")) Then
		response.redirect "rentalcategories.asp"
	Else 
		iReservationTempId = CLng(request("rti"))
	End If 

	' pull the data out of the holding table
	bHasData = SetPageVariables( iReservationTempId, iOrgId )

	If bHasData Then 
		ClearTempReservation iReservationTempId, iOrgId
	Else
		' Take them somewhere safe, as their data is gone.
		response.redirect "rentalcategories.asp"
	End If 
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

	<script src="../prototype/prototype-1.6.0.2.js"></script>
	<script src="../scriptaculous/src/scriptaculous.js"></script>

	<script src="../scripts/isvaliddate.js"></script>

	<script>
	<!--

		var doCalendar = function( sField ) {
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
		};

		var displayScreenMsg = function(iMsg) {
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		};

		var clearScreenMsg = function() {
			$("screenMsg").innerHTML = "&nbsp;";
		};

		var viewSingleDate = function() {
			// set the view type
			$("frmviewtype").value = "viewsingledate";

			// make sure the date field is filled in
			if ($("selectdate").value == "")
			{
				displayScreenMsg("Please enter a date, then try viewing again.");
				$("selectdate").focus();
				return false;
			}

			// blank out the other dates
			$("startdate").value = "";
			$("enddate").value = "";

			// submit the form
			document.frmRentalSearch.submit();
		};

		var viewSelectedDays = function() {
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

			if ($("enddate").value == "")
			{
				displayScreenMsg("Please enter an end date, then try viewing again.");
				$("enddate").focus();
				return false;
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
		};

		var SelectDate = function( iRentalId, sDate ) {
			//alert( sDate );
			$("selectedrid").value = iRentalId;
			$("selecteddate").value = sDate;
			document.selectForm.submit();
		};

	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

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
				<br />
				<input type="text" id="selectdate" name="selectdate" readonly="readonly" value="<%=sSelectDate%>" size="10" maxlength="10" onclick="javascript:void doCalendar('selectdate');" />
				&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('selectdate');" /></span>
				<br />
				<br />
				<input type="button" class="button" name="viewsingledate" value="View Single Date" onclick="viewSingleDate();" />
			</td>
			<td id="orcolumn" align="center" valign="middle">
				OR
			</td>
			<td class="selecttitle" align="center" colspan="2">
				Pick a range of dates and the<br />
				days of the week you wish to view.
				<table>
					<tr>
						<td align="center" nowrap="nowrap">
							<span class="respCol">
							Start Date: 
							</span>
							<span class="respCol">
							<input type="text" id="startdate" name="startdate" readonly="readonly" value="<%=sStartDate%>" size="10" maxlength="10" onclick="javascript:void doCalendar('startdate');" />
							&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
							</span>
						</td>
						<td align="center" nowrap="nowrap">
							<span class="respCol">
							End Date: 
							</span>
							<span class="respCol">
							<input type="text" id="enddate" name="enddate" readonly="readonly" value="<%=sEndDate%>" size="10" maxlength="10" onclick="javascript:void doCalendar('enddate');" />
							&nbsp;<span class="calendarimg"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
							</span>

						</td>
					</tr>
					<tr>
						<td colspan="2" align="center" nowrap="nowrap">
							<span class="respCol">
							<input type="checkbox" name="weeklydow" value="1" <%=s1Checked%> />Su
							<input type="checkbox" name="weeklydow" value="2" <%=s2Checked%> />Mo
							<input type="checkbox" name="weeklydow" value="3" <%=s3Checked%> />Tu
							<input type="checkbox" name="weeklydow" value="4" <%=s4Checked%> />We
							</span>
							<span class="respCol">
							<input type="checkbox" name="weeklydow" value="5" <%=s5Checked%> />Th
							<input type="checkbox" name="weeklydow" value="6" <%=s6Checked%> />Fr
							<input type="checkbox" name="weeklydow" value="7" <%=s7Checked%> />Sa
							</span>
						</td>
					</tr>
				</table>
				<input type="button" class="button" name="viewselecteddays" value="View Selected Days" onclick="viewSelectedDays();" />
			</td>
		</tr>

	</table>
</form>

<%
	If sViewType <> "none" and (sSelectDate <> "" or sStartDate <> "") Then
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

'		For x = 0 To UBound(aWantedDates, 2) 
'			response.write aWantedDates(0,x) & " &mdash; " & aWantedDates(1,x) & "<br /><br />"
'		Next 

		' They have pressed a button, so show some results
		ShowRentalAvailability iCategoryId, iRentalId, aWantedDates
	'elseIf sViewType <> "none" and sStartDate = "" Then
	elseif sViewType <> "none" and sSelectDate = "" and sStartDate = "" Then
		response.redirect "rentalcalendar.asp?rid=" & iRentalId
	

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

<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

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

	sSql = "SELECT R.rentalid, rentalname, locationname, ISNULL(width,'') AS width, ISNULL(length,'') AS length, "
	sSql = sSql & " ISNULL(capacity,'') AS capacity, ISNULL(shortdescription,'') AS shortdescription, nocosttorent, "
	sSql = sSql & " ISNULL(iconimageurl,'') AS iconimageurl, publiccanreserve  "
	sSql = sSql & " FROM egov_rentals_list R, egov_rentals_to_categories C "
	sSql = sSql & " WHERE R.rentalid = C.rentalid AND R.publiccanview = 1 AND R.publiccanreserve = 1 "
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
			response.write "<td valign=""top"" class=""iconcell"">" 
			' Show the name
			'response.write vbcrlf & "<div class=""rentalschedule"">"
			If oRs("iconimageurl") <> "" Then 
				response.write "<img src=""" & replace(oRs("iconimageurl"),"http://www.egovlink.com","") & """ alt=""" & oRs("rentalname") & """ title=""" & oRs("rentalname") & """ class=""availabilityimg"" />"
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			response.write "<td colspan=""2"" valign=""top"" align=""left"" class=""desccolumn"">"
			response.write "<p><span class=""schedulerentalname"">"
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
				If iCount Mod 2 = 0 Then
					sClass = " class=""altrow"" "
				Else
					sClass = ""
				End If 
				response.write "<tr" & sClass & ">"
				response.write "<td valign=""top"" align=""left"" class=""datecolumn"">"
				response.write "<span class=""datedisplay"">" & DateValue(CDate(aWantedDates(0,x))) & " &nbsp;&nbsp; " & WeekDayName(Weekday(CDate(aWantedDates(0,x)))) & "</span>"
				response.write "</td>"
				response.write "<td valign=""top"" align=""left"" class=""datecolumn"" colspan=""2""><span class=""schedulerentalname"">Availability</span></td></tr>"

				response.write "<tr" & sClass & ">"
				response.write "<td>&nbsp;</td>"
				response.write "<td valign=""top"" align=""left"" class=""availabledatecolumn"" nowrap=""nowrap"">"
				bOffSeasonFlag = GetOffSeasonFlag( oRs("rentalid"), CDate(aWantedDates(0,x)) )
				'response.write "bOffSeasonFlag = " & bOffSeasonFlag & "<br /><br />"

				bHasHours = DisplayAvailability( oRs("rentalid"), CDate(aWantedDates(0,x)), bOffSeasonFlag )
				
				'response.write "</td>"

				'response.write "<td valign=""top"" align=""left"" class=""selecttime"">"
				If bHasHours Then 
					response.write "<input type=""button"" class=""button selecttime"" value=""Select this Date and Continue"" onclick=""SelectDate( " & oRs("rentalid") & ", '" & aWantedDates(0,x) & "' );"" />"
				'Else
				'	response.write "&nbsp;"
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
' boolean SetPageVariables( iReservationTempId, iOrgId )
'--------------------------------------------------------------------------------------------------
Function SetPageVariables( ByVal iReservationTempId, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT cid, rid, viewtype, rentalid,  selecteddate, selectdate, startdate, enddate, wanteddows "
	sSql = sSql & " FROM egov_rentalreservationstemppublic "
	sSql = sSql & " WHERE reservationtempid = " & iReservationTempId
	sSql = sSql & " AND orgid = " & iOrgId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iCategoryId = CLng(oRs("cid"))
		iRentalId = CLng(oRs("rid"))
		sViewType = oRs("viewtype")
		sWantedDOWs = oRs("wanteddows")
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
