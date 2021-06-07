<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaldayedit.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Edit the daily schedule for each rental. Called from rentaledit.asp
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iDayId, sRentalName, sLocationName, sSeason, sDayName, sIsAvailableToPublic, sOpeningHour
Dim sOpeningMinute, sOpeningAmPm, sClosingHour, sClosingMinute, sClosingAmPm, sClosingDay, iRentalId
Dim sLatestStartHour, sLatestStartMinute, sLatestStartAmPm, sMinimumRental, sMinimumRentalTimeTypeId
Dim sPostBuffer, sPostBufferTimeTypeId, iMaxRateRows, sLoadMsg, bOrgHasAccounts, sIsOpen
Dim bShowCopyAnother

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "create edit rentals", sLevel	' In common.asp

iDayId = CLng(request("dayid"))
sRentalName = ""
sLocationName = ""
sSeason = ""
sDayName = ""
sIsAvailableToPublic = ""
sOpeningHour = 0
sOpeningMinute = 0
sOpeningAmPm = "AM"
sClosingHour = 0
sClosingMinute = 0
sClosingAmPm = "PM"
sClosingDay = 0
iRentalId = 0
sLatestStartHour = 0
sLatestStartMinute = 0
sLatestStartAmPm = "PM"
sMinimumRental = ""
sMinimumRentalTimeTypeId = 0
sPostBuffer = ""
sPostBufferTimeTypeId = 0
iMaxRateRows = 0
sIsOpen = ""
bShowCopyAnother = False 

GetDayInfo iDayId

If iRentalId = CLng(0) Then
	' If they are trying to see something that is not in their org then take them away.
	response.redirect "rentalslist.asp"
End If 

If request("s") = "u" Then
	sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
End If 
If request("s") = "c" Then
	sLoadMsg = "displayScreenMsg('The Copy Was Successful');"
	bShowCopyAnother = True 
End If 

' Not every org has general ledger accounts so we need to be able to hide/show accordingly.
bOrgHasAccounts = OrgHasFeature("gl accounts")


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>


	<script language="Javascript">
	<!--

		function ValidatePrice( oPrice )
		{
			// Remove any extra spaces
			oPrice.value = removeSpaces(oPrice.value);
			//Remove commas that would cause problems in validation
			oPrice.value = removeCommas(oPrice.value);

			// Validate the format of the price
			if (oPrice.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oPrice.value);
				if ( Ok )
				{
					oPrice.value = format_number(Number(oPrice.value),2);
				}
				else 
				{
					oPrice.value = "";
					alert("Rates must be numbers in currency format or blank.\nPlease correct to continue.");
					setfocus(oPrice);
					return false;
				}
			}
		}

		function ValidateTimeAmount( oTime )
		{
			// Validate the time amounts
			if (oTime.value != '')
			{
				// Remove any extra spaces
				oTime.value = removeSpaces(oTime.value);
				//Remove commas that would cause problems in validation
				oTime.value = removeCommas(oTime.value);

				rege = /^\d*$/;
				Ok = rege.test(oTime.value);
				if ( ! Ok )
				{
					oTime.value = "";
					alert("This field must be a positive integer.\nPlease correct this and try saving again.");
					setfocus(oTime);
					return false;
				}
			}
		}

		function Validate()
		{
			var iPriceTypeId;
			//check that any selected rate has an amount
			for (t = 1; t <= parseInt(document.frmRentalRates.maxraterows.value); t++)
			{
				if ($("pricetypeid" + t).checked == true)
				{
					iPriceTypeId = $("pricetypeid" + t).value;
					if ($("amount" + iPriceTypeId).value == "")
					{
						alert("You have selected a rate but not input an amount.\nPlease correct this and try saving again.");
						$("amount" + iPriceTypeId).focus();
						return false;
					}
				}
			}

			//alert('OK to go');
			// submit the page
			document.frmRentalRates.submit();
		}

		function SetUpPage()
		{
			<%=sLoadMsg%>
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
			$("screenMsg").innerHTML = "";
		}

	//-->
	</script>

</head>

<body onload="SetUpPage();">

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Edit Rental Rates</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<table id="screenMsgtable"><tr><td>
				<span id="screenMsg"></span>
				<input type="button" class="button" value="<< Back" onclick="location.href='rentaledit.asp?rentalid=<%=iRentalId%>';" />
<%			If bShowCopyAnother Then		%>
					&nbsp;<input type="button" class="button" value="Copy Another Day" onclick="location.href='rentaldaycopy.asp?rentalid=<%=iRentalId%>';" />
<%			End If		%>
			</td></tr></table>

			<p class="ratesubtitles">
				<%= sRentalName & " &ndash; " & sLocationName %><br />
			</p>

			<p class="ratedaysubtitles">
				<%= sDayName & " &ndash; " & sSeason %><br />
			</p>

			<form name="frmRentalRates" method="post" action="rentaldayupdate.asp">
				<input type="hidden" name="dayid" value="<%=iDayId%>" />
				<input type="hidden" name="rentalid" value="<%=iRentalId%>" />
				<p>
					<input type="checkbox" name="isopen" id="isopen" <%=sIsOpen%>/>
					&nbsp; Is open on this day of the week
				</p>
				<p>
					<input type="checkbox" name="isavailabletopublic" id="isavailabletopublic" <%=sIsAvailableToPublic%>/>
					&nbsp; Is available for public side reservations on this day of the week
				</p>
				<p>
					Opens at: <% ShowHourPicks "openinghour", sOpeningHour, ""	%>&nbsp;:&nbsp;
					<% ShowMinutePicks "openingminute", sOpeningMinute, ""	%>&nbsp;
<%					ShowAmPmPicks "openingampm", sOpeningAmPm, ""	%>
				</p>
				<p>
					Closes at: <% ShowHourPicks "closinghour", sClosingHour, ""	%>&nbsp;:&nbsp;
					<% ShowMinutePicks "closingminute", sClosingMinute, ""%>&nbsp;
<%					ShowAmPmPicks "closingampm", sClosingAmPm, ""	%> &nbsp; <% ShowSameNextDayPick "closingday", sClosingDay, ""	%>
				</p>
				<p>
					Latest Reservation Start Time: <% ShowHourPicks "lateststarthour", sLatestStartHour, ""	%>&nbsp;:&nbsp;
					<% ShowMinutePicks "lateststartminute", sLatestStartMinute, "" %>&nbsp;
<%					ShowAmPmPicks "lateststartampm", sLatestStartAmPm, ""	%>
				</p>
				<p>
					Time Buffer After a Reservation: <input type="text" id="postbuffer" name="postbuffer" value="<%=sPostBuffer%>" size="3" maxlength="3" onchange="ValidateTimeAmount( this );" /> 
					<% ShowTimeTypePicks "postbuffertimetypeid", sPostBufferTimeTypeId, True 	%>
				</p>

				<p>
					Minimum Rental Time: <input type="text" id="minimumrental" name="minimumrental" value="<%=sMinimumRental%>" size="3" maxlength="3" onchange="ValidateTimeAmount( this );" />
					<% ShowTimeTypePicks "minimumrentaltimetypeid", sMinimumRentalTimeTypeId, False 	%>
				</p>
				<strong>Rates: (If there is no cost to rent this, select &quot;Everyone&quot; and set the rate to $0 per hour.)</strong><br /><br />
				<table id="rentalratestable" border="0" cellpadding="0" cellspacing="0">
					<tr><th>Type</th><th>
<%						If bOrgHasAccounts Then 										
							response.write "Account"
						Else
							response.write "&nbsp;"
						End If 
%>
						</th><th>Rate<br />Type</th><th>Rate</th><th>Starts At</th></tr>
<%					iMaxRateRows = ShowRentalRates( iDayId, iRentalId, bOrgHasAccounts )		%>
				</table>
				<input type="hidden" id="maxraterows" name="maxraterows" value="<%=iMaxRateRows%>" />
				<p>
					<input type="button" class="button" value="Save Changes" onclick="Validate();" />
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
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' GetDayInfo iDayId 
'--------------------------------------------------------------------------------------------------
Sub GetDayInfo( ByVal iDayId )
	Dim sSql, oRs

	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, D.weekdayname, D.isoffseason, D.isavailabletopublic, D.isopen, "
	sSql = sSql & "ISNULL(D.openinghour,0) AS openinghour, ISNULL(D.openingminute,0) AS openingminute, "
	sSql = sSql & "ISNULL(D.openingampm,'AM') AS openingampm, ISNULL(D.closinghour,0) AS closinghour, "
	sSql = sSql & "ISNULL(D.closingminute,0) AS closingminute, ISNULL(D.closingampm,'PM') AS closingampm, "
	sSql = sSql & "ISNULL(D.closingday,0) AS closingday, ISNULL(D.lateststarthour,0) AS lateststarthour, "
	sSql = sSql & "ISNULL(D.lateststartminute,0) AS lateststartminute, ISNULL(D.lateststartampm,'PM') AS lateststartampm, "
	sSql = sSql & "ISNULL(D.postbuffer,0) AS postbuffer, ISNULL(D.postbuffertimetypeid,0) AS postbuffertimetypeid, "
	sSql = sSql & "ISNULL(D.minimumrental,0) AS minimumrental, ISNULL(D.minimumrentaltimetypeid,0) AS minimumrentaltimetypeid "
	sSql = sSql & "FROM egov_rentaldays D, egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE D.rentalid = R.rentalid AND R.locationid = L.locationid AND D.dayid = " & iDayId & " AND R.orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iRentalId = oRs("rentalid")
		sRentalName = oRs("rentalname")
		sLocationName = oRs("locationname")
		If oRs("isoffseason") Then 
			sSeason = "Off Season"
		Else
			sSeason = "In Season"
		End If 
		sDayName = oRs("weekdayname")
		If oRs("isavailabletopublic") Then 
			sIsAvailableToPublic = " checked=""checked"" "
		Else
			sIsAvailableToPublic = ""
		End If 
		sOpeningHour = oRs("openinghour")
		sOpeningMinute = oRs("openingminute")
		sOpeningAmPm = oRs("openingampm")
		sClosingHour = oRs("closinghour")
		sClosingMinute = oRs("closingminute")
		sClosingAmPm = oRs("closingampm")
		sClosingDay = oRs("closingday")
		sLatestStartHour = oRs("lateststarthour")
		sLatestStartMinute = oRs("lateststartminute")
		sLatestStartAmPm = oRs("lateststartampm")
		If clng(oRs("minimumrental")) > clng(0) Then
			sMinimumRental = oRs("minimumrental")
		Else
			sMinimumRental = ""
		End If 
		sMinimumRentalTimeTypeId = oRs("minimumrentaltimetypeid")
		If clng(oRs("postbuffer")) > clng(0) Then
			sPostBuffer = clng(oRs("postbuffer"))
		Else
			sPostBuffer = ""
		End If 
		sPostBufferTimeTypeId = oRs("postbuffertimetypeid")
		If oRs("isopen") then
			sIsOpen = " checked=""checked"" "
		Else
			sIsOpen = ""
		End If 
	Else
		iRentalId = CLng(0)
	End If
	
	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' integer iRateCount = ShowRentalRates( iDayId, iRentalId, bOrgHasAccounts )
'--------------------------------------------------------------------------------------------------
Function ShowRentalRates( ByVal iDayId, ByVal iRentalId, ByVal bOrgHasAccounts )
	Dim oRs, sSql, iAccountNo, iRateTypeId, sRateAmount, iStartHour, iStartMinute, sStartAmPm, iRowCount

	iRowCount = 0
	sSql = "SELECT pricetypeid, pricetypename, hasstarttime, isbaseprice, isfee "
	sSql = sSql & " FROM egov_price_types WHERE isforrentals = 1 AND isrentalflatfee = 0 AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		sSelected = GetDayRateInfo( iDayId, iRentalId, oRs("pricetypeid"), iAccountNo, iRateTypeId, sRateAmount, iStartHour, iStartMinute, sStartAmPm )
		response.write vbcrlf & "<tr>"
		' Show the Price Type
		response.write "<td class=""type""><input type=""checkbox"" id=""pricetypeid" & iRowCount & """ name=""pricetypeid"" value=""" & oRs("pricetypeid") & """" & sSelected & " /> &nbsp;"
		response.write oRs("pricetypename") & " ("
		If oRs("isbaseprice") Then
			response.write "base"
		Else
			If oRs("isfee") Then
				response.write "fee"
			Else
				response.write "+"
			End If 
		End If 
		response.write ")</td>"
		
		' Show the account
		response.write "<td align=""center"">"
		If bOrgHasAccounts Then 
			ShowAccountPicks "accountid" & oRs("pricetypeid"), iAccountNo, False 
		Else
			response.write "<input type=""hidden"" id=""accountid" &  oRs("pricetypeid") & """ name=""accountid" &  oRs("pricetypeid") & """ value=""0"" />"
		End If 
		response.write "</td>"

		' Show the rate type
		response.write "<td align=""center"">"
		ShowRateTypePicks "ratetypeid" & oRs("pricetypeid"), iRateTypeId
		response.write "</td>"

		' Show the rate
		response.write "<td align=""center"">"
		response.write "<input type=""text"" id=""amount" & oRs("pricetypeid") & """ name=""amount" & oRs("pricetypeid") & """ value=""" & sRateAmount & """ size=""7"" maxlength=""7"" onchange=""ValidatePrice(this);"" />"
		response.write "</td>"

		' Show the starts at time
		response.write "<td align=""center"">"
		response.write "<input type=""hidden"" name=""hasstarttime"" value=""" & oRs("hasstarttime") & """ />"
		If oRs("hasstarttime") Then
			ShowHourPicks "starthour" & oRs("pricetypeid"), iStartHour, ""
			response.write ":"
			ShowMinutePicks "startminute" & oRs("pricetypeid"), iStartMinute, ""
			'response.write "&nbsp;"
			ShowAmPmPicks "startampm" & oRs("pricetypeid"), sStartAmPm, ""
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"

		response.write "</tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	ShowRentalRates = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' string sCheckStatus = GetDayRateInfo( iDayId, iRentalId, iPriceTypeId, iAccountNo, iRateTypeId, sRateAmount, iStartHour, iStartMinute, sStartAmPm )
'--------------------------------------------------------------------------------------------------
Function GetDayRateInfo( ByVal iDayId, ByVal iRentalId, ByVal iPriceTypeId, ByRef iAccountNo, ByRef iRateTypeId, ByRef sRateAmount, ByRef iStartHour, ByRef iStartMinute, ByRef sStartAmPm )
	Dim oRs, sSql

	sSql = "SELECT ISNULL(accountid,0) AS accountid, ISNULL(ratetypeid,0) AS ratetypeid, ISNULL(amount,0) AS amount, "
	sSql = sSql & " ISNULL(starthour,0) AS starthour, ISNULL(startminute,0) AS startminute, ISNULL(startampm,'AM') AS startampm "
	sSql = sSql & " FROM egov_rentaldayrates "
	sSql = sSql & " WHERE rentalid = " & iRentalId & " AND dayid = " & iDayId & " AND pricetypeid = " & iPriceTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		iAccountNo = oRs("accountid")
		iRateTypeId = oRs("ratetypeid")
		'If CDbl(oRs("amount")) > CDbl(0.00) Then 
			sRateAmount = FormatNumber(oRs("amount"),2,,,0)
		'Else
		'	sRateAmount = ""
		'End If 
		iStartHour = oRs("starthour")
		iStartMinute = oRs("startminute")
		sStartAmPm = oRs("startampm")
		GetDayRateInfo = " checked=""checked"" "
	Else
		iAccountNo = 0
		iRateTypeId = 0
		sRateAmount = ""
		iStartHour = 0
		iStartMinute = 0
		sStartAmPm = "AM"
		GetDayRateInfo = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' ShowRateTypePicks sSelectName, iRateTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowRateTypePicks( ByVal sSelectName, ByVal iRateTypeId )
	Dim sSql, oRs

	sSql = "SELECT ratetypeid, ratetype FROM egov_rentalratetypes WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""" & sSelectName & """ name=""" & sSelectName & """>"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("ratetypeid") & """"
		If CLng(oRS("ratetypeid")) = CLng(iRateTypeId) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("ratetype") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing 
End Sub 




%>
