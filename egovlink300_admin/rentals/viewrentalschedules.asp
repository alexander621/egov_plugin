<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewrentalschedules.asp
' AUTHOR: Steve Loar
' CREATED: 08/13/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  List of rentals. From here you can create or edit rentals
'
' MODIFICATION HISTORY
' 1.0   08/13/2009	Steve Loar - INITIAL VERSION
' 1.1	05/11/2010	Steve Loar - Modified mechanics of date selections
' 1.2	10/12/2010	Steve Loar - Adding purpose to show on internals in addition to the blocks.
' 1.3	03/24/2011	Steve Loar - hide deactivated rentals
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sSelect, sFrom, fromDate, toDate, sShowItems, iRentalId, bShowItems, sRentalPageBreak, sShowRenterName
Dim bRentalPageBreak, iSupervisorUserId, sShowIsOnHold, bShowIsOnHold, bShowRenterName

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "viewrentalschedules", sLevel	' In common.asp

sSearch = ""
sSelect = ""

fromDate = Request("fromDate")
toDate = Request("toDate")

' IF EMPTY Dates then DEFAULT TO This Week
If toDate = "" or IsNull(toDate) Then
	toDate = Date() 
	If Weekday(toDate) < 7 Then 
		' if not Saturday, set the start to Saturday
		toDate = DateAdd("d", (7 - Weekday(date)), date)
	End If 
End If

If fromDate = "" or IsNull(fromDate) Then 
	fromDate = Date()
	If Weekday(fromDate) > 1 Then 
		' if not Sunday, set the end to that
		fromDate = DateAdd("d", -(Weekday(date) - 1), date)
	End If 
End If

If request("showitems") = "on" Then
	sShowItems = "checked=""checked"" "
	bShowItems = True 
Else 
	sShowItems = ""
	bShowItems = False 
End If 

If request("showisonhold") = "on" Then
	sShowIsOnHold = "checked=""checked"" "
	bShowIsOnHold = True 
Else
	sShowIsOnHold = ""
	bShowIsOnHold = False 
End If 

If request("rentalpagebreak") = "on" Then
	sRentalPageBreak = "checked=""checked"" "
	bRentalPageBreak = True 
Else 
	If UCase(request.ServerVariables("REQUEST_METHOD")) = "POST" Then 
		sRentalPageBreak = ""
		bRentalPageBreak = False 
	Else
		sRentalPageBreak = "checked=""checked"" "
		bRentalPageBreak = True 
	End If 
End If 

if request("showrentername") = "on" then
  sShowRenterName = "checked=""checked"" "
  bShowRenterName = true
else
  if ucase(request.servervariables("request_method")) = "POST" then
     sShowRenterName = ""
     bShowRenterName = false
  else
     sShowRenterName = "checked=""checked"" "
     bShowRenterName = true
  end if
end if

If request("rentalid") <> "" Then
	iRentalId = request("rentalid")
	If iRentalId <> "0" Then 
		sRentalIdType = Left(iRentalId, 1)
		iActualId = Mid(iRentalId, 2)
		If sRentalIdType = "R" Then
			sSearch = sSearch & " AND R.rentalid = " & iActualId
		Else
			sSearch = sSearch & " AND L.locationid = " & iActualId
		End If 
	End If 
Else
	' Find their first real pick and default to that
	iRentalId = GetFirstRentalIdPick()
	sSearch = sSearch & " AND R.rentalid = " & Mid(iRentalId, 2)
End If 


If request("supervisoruserid") <> "" Then
	iSupervisorUserId = CLng(request("supervisoruserid"))
	If iSupervisorUserId > CLng(0) Then
		sSearch = sSearch & " AND R.supervisoruserid = " & iSupervisorUserId
	End If 
Else
	iSupervisorUserId = 0
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>

	<script language="Javascript">
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

		function doCalendar( sField ) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			var sSelectedDate = $(sField).value;

			// Set the to date to the from date
			if (sField == "toDate")
			{
				sSelectedDate = $("fromDate").value;
				//alert( $("fromDate").value );
			}

			//alert( sField + ": " + sSelectedDate );

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmReservationSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function RefreshResults()
		{
			document.frmReservationSearch.submit();
		}

		function callSearch( sDate, sRental, iCategory, iLocation )
		{
			$("startdate").value = sDate;
			$("enddate").value = sDate;
			$("rentalname").value = sRental;
			$("recreationcategoryid").value = iCategory;
			$("locationid").value = iLocation;
			document.frmsearch.submit();
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
'		<input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
'		<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />&nbsp;&nbsp;
%>
	<%	If request("rt") = "r" Then %>
			&nbsp;&nbsp;<input type="button" class="button" value="<< Back To Reservation" onclick="location.href='reservationedit.asp?reservationid=<%=iReservationId%>';" />	
	<%	End If	%>
	</div>

<%
'	<object id="factory" viewastext  style="display:none"
'	   classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
'	   codebase="../includes/smsx.cab#Version=6,3,434,12">
'	</object>
%>
	<!--END: THIRD PARTY PRINT CONTROL-->

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Rental Schedule</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmReservationSearch" method="post" action="viewrentalschedules.asp">
							<table id="scheduleselections" cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Date Range:</td>
									<td>
										From:
										<input type="text" id="fromDate" name="fromDate" value="<%=fromDate%>" readonly="readonly" size="10" maxlength="10" onclick="javascript:void doCalendar('fromDate');" />
										<a href="javascript:void doCalendar('fromDate');"><img src="../images/calendar.gif" border="0" /></a>
										&nbsp; To:
										<input type="text" id="toDate" name="toDate" value="<%=toDate%>" readonly="readonly" size="10" maxlength="10" onclick="javascript:void doCalendar('toDate');" />
										<a href="javascript:void doCalendar('toDate');"><img src="../images/calendar.gif" border="0" /></a>
										&nbsp;
										<%DrawDateChoices "Date" %>
									</td>
								</tr>
								<tr>
									<td>Rental:</td><td><% ShowRentalLocationPicks iRentalId, True, True 	' In rentalsguifunctions.asp %></td>
								</tr>
								<tr>
									<td>Supervisor:</td><td><% ShowRentalSupervisors iSupervisorUserId, "All Supervisors" 	' In rentalsguifunctions.asp %></td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td valign="middle"><input type="checkbox" name="showitems" <%=sShowItems%> /> Show the items needed for each reservation.
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td valign="middle"><input type="checkbox" name="showisonhold" <%=sShowIsOnHold%> /> Include reservations that are On Hold.
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td valign="middle"><input type="checkbox" name="rentalpagebreak" <%=sRentalPageBreak%> /> Print rentals on separate pages.
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td valign="middle"><input type="checkbox" name="showrentername" <%=sShowRenterName%> /> Show Renter Name in results.
									</td>
								</tr>
								<tr>
			    					<td colspan="2"><input class="button" type="button" value="Refresh Results" onclick="RefreshResults();" /></td>
  								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->


<%				
			ShowReservations sSearch, fromDate, toDate, bShowItems, bRentalPageBreak, bShowIsOnHold, bShowRenterName
%>			

			<form name="frmsearch" method="post" action="rentalsearch.asp">
				<input type="hidden" id="startdate" name="startdate" value="" />
				<input type="hidden" id="enddate" name="enddate" value="" />
				<input type="hidden" id="rentalname" name="rentalname" value="" />
				<input type="hidden" id="recreationcategoryid" name="recreationcategoryid" value="" />
				<input type="hidden" name="starthour" value="1" />
				<input type="hidden" name="startminute" value="00" />
				<input type="hidden" name="startampm" value="PM" />
				<input type="hidden" name="endhour" value="2" />
				<input type="hidden" name="endminute" value="00" />
				<input type="hidden" name="endampm" value="PM" />
				<input type="hidden" name="endday" value="0" />
				<input type="hidden" name="occurs" value="o" />
				<input type="hidden" name="orderby" value="1" />
				<input type="hidden" id="locationid" name="locationid" value="0" />
				<input type="hidden" name="periodtypeid" value="<%=GetIsSelectedPeriodTypeId()%>" />
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
' void ShowReservations sSearch, fromDate, toDate, bShowItems, bRentalPageBreak, bShowIsOnHold
'--------------------------------------------------------------------------------------------------
Sub ShowReservations( ByVal sSearch, ByVal fromDate, ByVal toDate, ByVal bShowItems, ByVal bRentalPageBreak, ByVal bShowIsOnHold, ByVal bShowRenterName )
	Dim dDate, x, iDateCount, sSql, oRs, iRentalsCount, sOpeningTime, sClosingTime

	iRentalsCount = 0

	' Get the rentals involved 
	sSql = "SELECT R.rentalid, R.rentalname, L.name AS locationname, L.locationid "
	sSql = sSql & " FROM egov_rentals R, egov_class_location L "
	sSql = sSql & " WHERE R.locationid = L.locationid AND R.isdeactivated = 0 AND R.orgid = " & session("orgid") & sSearch
	sSql = sSql & " ORDER BY L.name, R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	' Loop through the rentals 
	Do While Not oRs.EOF
		iRentalsCount = iRentalsCount + 1
		If iRentalsCount > 1 Then
			response.write vbcrlf & "<div"
			If bRentalPageBreak Then 
				' This class only forces the page break. Found only in receiptprint.css
				response.write " class=""rentalsstart"""
			End If 
			response.write ">"
			response.write vbcrlf & "<p>"
			response.write vbcrlf & "<font size=""+1""><strong>Rental Schedule</strong></font><br />"
			response.write vbcrlf & "</p>"
			response.write vbcrlf & "</div>"
		End If 
		' Show the name
		response.write vbcrlf & "<div class=""rentalschedule""><span class=""schedulerentalname"">" & oRs("rentalname") & " &ndash; " & oRs("locationname") & "</span>"

		' Get a recreationcategoryid for the "make a reservation" button
		iRecreationCategoryId = GetARecreationCategoryId( oRs("rentalid") )

		' Reset the date counter
		iDateCount = DateDiff("d", CDate(fromDate), CDate(toDate))

		' Loop through the date range
		For x = 0 To iDateCount
			dDate = DateAdd("d", x, CDate(fromDate))
			response.write vbcrlf & "<div class=""showdate"">"
			If bShowItems Then
				response.write vbcrlf & "<span class=""itemtitle"">Items</span>"
			End If 
			response.write vbcrlf & "<span class=""showdate"">" & dDate & "</span> <span class=""showdayname"">" & WeekDayName(Weekday(dDate)) & "</span> &nbsp; "
			bOpenOnDate = RentalIsOpenOnThisDate( oRs("rentalid"), DateValue(dDate) )
			If bOpenOnDate Then 
				response.write "<input type=""button"" class=""button"" value=""Make a Reservation"" onclick=""callSearch( '" & dDate & "', '" & FormatForJavaScript( oRs("rentalname") ) & "', " & iRecreationCategoryId & ", " & oRs("locationid") & " )"" />" ' FormatForJavaScript( ) is in common.asp
			End If 
			response.write vbcrlf & "</div>"
			response.write vbcrlf & "<div class=""rentaltimes"">"

			If bOpenOnDate Then 
				GetRentalOpeningClosingTime oRs("rentalid"), DateValue(dDate), sOpeningTime, sClosingTime
			Else
				sOpeningTime = "Closed on this date"
				sClosingTime = ""
			End If 

			' Show the rental's reservations for this date
			ShowRentalReservations oRs("rentalid"), DateValue(dDate), bShowItems, sOpeningTime, sClosingTime, bShowIsOnHold, bShowRenterName

			response.write vbcrlf & "</div>"
		Next 
		response.write vbcrlf & "</div>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRentalReservations iRentalId, dStartDate, bShowItems, sOpeningTime, sClosingTime, bShowIsOnHold
'--------------------------------------------------------------------------------------------------
Sub ShowRentalReservations( ByVal iRentalId, ByVal dStartDate, ByVal bShowItems, sOpeningTime, ByVal sClosingTime, ByVal bShowIsOnHold, ByVal bShowRenterName )
	Dim sSql, oRs, dReservationStartTime, dReservationEndTime, sStartHour, sStartMinute, sStartAmPm
	Dim sEndHour, sEndMinute, sEndAmPm, sActivityNo, sClassName, sShowIsOnHold

	dReservationStartTime = dStartDate & " 00:00:00 AM"
	dReservationEndTime = DateAdd("d", 1, CDate(dReservationStartTime))

	If bShowIsOnHold Then
		sShowIsOnHold = ""
	Else
		sShowIsOnHold = " AND R.isonhold = 0 "
	End If 

	sSql = "SELECT R.reservationid, D.reservationdateid, D.reservationstarttime, D.billingendtime, "
	sSql = sSql & "ISNULL(R.totalamount, 0.00) AS totalamount, R.isonhold, "
	sSql = sSql & "(R.totalamount + R.totalrefunded - R.totalpaid + R.totalrefundfees) AS balancedue, "
	sSql = sSql & "R.rentaluserid, R.timeid, T.reservationtype, T.reservationtypeselector, R.purpose "
	sSql = sSql & " FROM egov_rentalreservationdates D, egov_rentalreservations R, egov_rentalreservationstatuses DS, "
	sSql = sSql & " egov_rentalreservationstatuses RS, egov_rentalreservationtypes T "
	sSql = sSql & " WHERE R.reservationid = D.reservationid AND D.rentalid = " & iRentalId
	sSql = sSql & " AND D.reservationstarttime BETWEEN '" & dReservationStartTime & "' AND '" & dReservationEndTime & "'"
	sSql = sSql & " AND D.statusid = DS.reservationstatusid AND DS.iscancelled = 0 "
	sSql = sSql & " AND R.reservationstatusid = RS.reservationstatusid AND RS.iscancelled = 0 "
	sSql = sSql & " AND R.reservationtypeid = T.reservationtypeid " & sShowIsOnHold
	sSql = sSql & " ORDER BY reservationstarttime"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<table class=""showtimes"" cellpadding=""2"" cellspacing=""0"" border=""0"">"

	If sOpeningTime <> "" Then
		If sOpeningTime <> "Closed on this date" Then 
			response.write vbcrlf & "<tr>"
			response.write "<td class=""timedisplay""><strong>" & sOpeningTime
			response.write "</strong></td>"
			response.write "<td class=""reservationtype"">Opens"
			response.write "</td>"
			response.write "<td colspan=""3"">&nbsp;</td>"
			response.write "</tr>"
		Else
			response.write vbcrlf & "<tr>"
			response.write "<td class=""timedisplay"">&nbsp;</td>"
			response.write "<td class=""reservationtype""><strong>Closed on this date</strong>"
			response.write "</td>"
			response.write "<td colspan=""3"">&nbsp;</td>"
			response.write "</tr>"
		End If 
	End If 

	If Not oRs.EOF Then 
		
		' Loop through the rentals 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			' From and to times
			response.write "<td class=""timedisplay""><strong>" & FormatTimeString( oRs("reservationstarttime") )
			response.write " &ndash; " & FormatTimeString( oRs("billingendtime") )
			response.write "</strong></td>"

			' Renter
			response.write "<td>"
			If oRs("reservationtypeselector") = "public" Then
    if bShowRenterName then
   				response.write GetCitizenName( oRs("rentaluserid") )
    else
       response.write "&nbsp;" & vbcrlf
    end if
			Else
				If oRs("reservationtypeselector") = "admin" Then
       if bShowRenterName then
			     		response.write GetAdminName( oRs("rentaluserid") )
       else
          response.write "&nbsp;"
       end if
				Else 
					If oRs("reservationtypeselector") = "class" Then
						sActivityNo = GetActivityNo( oRs("timeid") )	' In rentalscommonfunctions.asp 
						sClassName = GetClassName( oRs("timeid") )	' In rentalscommonfunctions.asp 
						response.write sActivityNo & "<br />"
						If Len(sClassName) > 27 Then
							response.write Left(sClassName,27) & "..."
						Else
							response.write sClassName
						End If 
					Else 
						response.write "&nbsp;"
					End If 
				End If 
			End If 
			response.write "</td>"

			' Reservation Type
			response.write "<td class=""reservationtype"" align=""center"">" & oRs("reservationtype")
			
			If oRs("isonhold") Then
				response.write "<br /><strong>On Hold</strong>"
			End If 
			response.write "</td>"

			' Balance Due
			response.write "<td class=""reservationtype"" align=""center"">" 
			If oRs("reservationtypeselector") = "public" Then
				If CDbl(oRs("balancedue")) > CDbl(0.00) Then 
					response.write "Owes " & FormatCurrency(oRs("balancedue"),2)
				Else
					If CDbl(oRs("totalamount")) > CDbl(0.00) Then 
						response.write "Paid"
					Else
						response.write "No Fees"
					End If 
				End If 
			Else
				If oRs("reservationtypeselector") = "block" Or oRs("reservationtypeselector") = "admin"  Then
					response.write oRs("purpose")
				Else 
					response.write "&nbsp;"
				End If 
				'response.write "&nbsp;"
			End If 
			response.write "</td>"

			' Reservation Edit button
			response.write "<td class=""reservationid"" align=""center""><input type=""button"" class=""button"" value=""Edit"" onclick=""location.href='reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" />"
			response.write "</td>"

			response.write "</tr>"

			If bShowItems And oRs("reservationtypeselector") <> "block" Then
				ShowItemsForDate oRs("reservationdateid")
			End If 

			oRs.MoveNext
		Loop
		
	Else 
		response.write "<tr><td class=""timedisplay"">&nbsp;</td><td colspan=""4"">Nothing Scheduled</td></tr>"
	End If 
	If sClosingTime <> "" Then
		response.write vbcrlf & "<tr>"
		response.write "<td class=""timedisplay""><strong>" & sClosingTime
		response.write "</strong></td>"
		response.write "<td class=""reservationtype"">Closes"
		response.write "</td>"
		response.write "<td colspan=""3"">&nbsp;</td>"
		response.write "</tr>"
	End If 
	response.write vbcrlf & "</table>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowItemsForDate iReservationDateId
'--------------------------------------------------------------------------------------------------
Sub ShowItemsForDate( ByVal iReservationDateId )
	Dim oRs, sSql, iItemCount

	iItemCount = clng(0)

	sSql = "SELECT reservationdateitemid, rentalitem, ISNULL(quantity,0) AS quantity "
	sSql = sSql & " FROM egov_rentalreservationdateitems "
	sSql = sSql & " WHERE reservationdateid = " & iReservationDateId
	sSql = sSql & " AND quantity IS NOT NULL "
	sSql = sSql & " ORDER BY rentalitem"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If clng(oRs("quantity")) > clng(0) Then 
			iItemCount = iItemCount + 1
			response.write vbcrlf & "<tr>"
			response.write "<td colspan=""4"" align=""right"">"
			response.write FormatNumber(oRs("quantity"),0,,,0) & "&nbsp;" & oRs("rentalitem")
			response.write "</td>"
			response.write "</tr>"
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 
	
	If clng(iItemCount) = clng(0) Then
		response.write "<tr><td colspan=""4"" align=""right"">No Items</td></tr>"
	End If 

End Sub 


'--------------------------------------------------------------------------------------------------
' string sRentalId = GetFirstRentalIdPick()
'--------------------------------------------------------------------------------------------------
Function GetFirstRentalIdPick()
	Dim sSql, oRs

	' The same query used to build the pick list from ShowRentalLocationPicks in rentalsguifunctions.asp
	sSql = "SELECT L.locationid , R.rentalid, L.name AS locationname, R.rentalname "
	sSql = sSql & "FROM egov_rentals R, egov_class_location L "
	sSql = sSql & "WHERE R.locationid = L.locationid AND R.orgid = " & session("orgid")
	sSql = sSql & "ORDER BY L.name, R.rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		' we want the first rental not the first location
		GetFirstRentalIdPick = "R" & oRs("rentalid")
	Else
		' hopefully this will not pull any records
		GetFirstRentalIdPick = "R0"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer iRecreationCategoryId = GetARecreationCategoryId( iRentalId )
'--------------------------------------------------------------------------------------------------
Function GetARecreationCategoryId( iRentalId )
	Dim sSql, oRs

	sSql = "SELECT recreationcategoryid FROM egov_rentals_to_categories WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetARecreationCategoryId = oRs("recreationcategoryid")
	Else
		GetARecreationCategoryId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer iPeriodTypeId = GetIsSelectedPeriodTypeId()
'--------------------------------------------------------------------------------------------------
Function GetIsSelectedPeriodTypeId()
	Dim sSql, oRs

	sSql = "SELECT periodtypeid FROM egov_rentalperiodtypes WHERE isselectedperiod = 1 AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetIsSelectedPeriodTypeId = oRs("periodtypeid")
	Else
		GetIsSelectedPeriodTypeId = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' void GetRentalOpeningClosingTime iRentalid, dDate, sOpeningTime, sClosingTime
'--------------------------------------------------------------------------------------------------
Sub GetRentalOpeningClosingTime( ByVal iRentalid, ByVal dDate, ByRef sOpeningTime, ByRef sClosingTime )
	Dim bOffSeasonFlag, sSql, oRs, iWeekDay

	bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(dDate)) )

	iWeekDay = Weekday(dDate)

	sSql = "SELECT isopen, isavailabletopublic, ISNULL(openinghour,0) AS openinghour, dbo.AddLeadingZeros(ISNULL(openingminute,0),2) AS openingminute, "
	sSql = sSql & " openingampm, closinghour, dbo.AddLeadingZeros(ISNULL(closingminute,0),2) AS closingminute, closingampm FROM egov_rentaldays "
	sSql = sSql & " WHERE rentalid = " & iRentalid & " AND orgid = " & session("orgid")
	sSql = sSql & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekDay

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isopen") Then 
			' valid hours are 1 to 12, so 0 means that there is not data for this date
			If clng(oRs("openinghour")) > clng(0) Then 
				sOpeningTime = oRs("openinghour") & ":" & oRs("openingminute") & " " & oRs("openingampm")
				sClosingTime = oRs("closinghour") & ":" & oRs("closingminute") & " " & oRs("closingampm")
			Else
				sOpeningTime = ""
				sClosingTime = ""
			End If 
		Else
			sOpeningTime = ""
			sClosingTime = ""
		End If 
	Else
		sOpeningTime = ""
		sClosingTime = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean RentalIsOpenOnThisDate( iRentalid, dDate )
'--------------------------------------------------------------------------------------------------
Function RentalIsOpenOnThisDate( ByVal iRentalid, ByVal dDate )
	Dim bOffSeasonFlag, sSql, oRs, iWeekDay

	bOffSeasonFlag = GetOffSeasonFlag( iRentalid, DateValue(CDate(dDate)) )

	iWeekDay = Weekday(dDate)

	sSql = "SELECT isopen FROM egov_rentaldays "
	sSql = sSql & " WHERE rentalid = " & iRentalid & " AND orgid = " & session("orgid")
	sSql = sSql & " AND isoffseason = " & bOffSeasonFlag & " AND dayofweek = " & iWeekDay

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isopen") Then 
			RentalIsOpenOnThisDate = True 
		Else
			RentalIsOpenOnThisDate = False
		End If 
	Else
		RentalIsOpenOnThisDate = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>


