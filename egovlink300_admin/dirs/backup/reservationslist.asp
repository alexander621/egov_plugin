<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: reservationslist.asp
' AUTHOR: Steve Loar
' CREATED: 03/28/2011
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This is a list rental reservations for a citizen
'
' MODIFICATION HISTORY
' 1.0   03/28/2011   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sName, iUserId, sReservedFromDate, sReservedToDate, sReservationFromDate, sReservationToDate

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "edit citizens", sLevel	' In common.asp

iUserId = CLng(request("u"))
sName   = GetCitizenName( iUserId )

If request("fromreserveddate") <> "" Then 
	sReservedFromDate = DateValue(CDate(request("fromreserveddate")))
Else
	sReservedFromDate = DateValue(DateSerial(Year(Now),1,1))
End If 

If request("toreserveddate") <> "" Then 
	sReservedToDate = DateValue(CDate(request("toreserveddate")))
Else
	sReservedToDate = DateValue(now)
End If 

If request("fromreservationdate") <> "" Then 
	sReservationFromDate = DateValue(CDate(request("fromreservationdate")))
Else
	sReservationFromDate = ""
End If 

If request("toreservationdate") <> "" Then 
	sReservationToDate = DateValue(CDate(request("toreservationdate")))
Else
	sReservationToDate = ""
End If 

iReservationTotal = CDbl(0.00)
iRefundTotal = CDbl(0.00)
iPaidTotal = CDbl(0.00)
iBalanceTotal = CDbl(0.00)

%>
<html lang="en">
<head>
	<meta charset="utf-8" />
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="./reservationliststyles.css" />

	<script type="text/javascript" src="https://code.jquery.com/jquery-1.5.min.js"></script>

	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="javascript">
	<!--

		function doCalendar( ToFrom ) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("../classes/calendarpicker.asp?p=1&updateform=reservation_list&updatefield=' + ToFrom + '", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function validate()
		{
			// validate the dates selected

			//submit the form
			document.reservation_list.submit();
		}

		$(document).ready(function() {
			//back button click
			$("#backbtn").click(function() {
				history.back();
			});

			//Search button click
			$("#searchbtn").click(function() {
				validate();
			});
		});

	//-->
	</script>
</head>

<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
	<div id="centercontent">
		<font size="+1"><strong>Rental Reservations of <%=sName%></strong></font><p>
		<input type="button" id="backbtn" value="<< Back" /><br />

		<fieldset>
			<legend>Search Criteria&nbsp;</legend><p>
			
				<form name="reservation_list" method="post" action="reservationslist.asp">
					<input type="hidden" name="u" value="<%=iUserId%>" />

					<table border="0" cellspacing="0" cellpadding="2">
					<tr>
						<td><strong>Reserved Date</strong></td>
						<td>From: </td>
						<td>
							<input type="text" id="fromreserveddate" name="fromreserveddate" value="<%=sReservedFromDate%>" size="10" maxlength="10">&nbsp;
							<a href="javascript:void doCalendar('fromreserveddate');"><img src="../images/calendar.gif" height="16" width="16" border="0"></a>
						</td>
						<td>To: </td>
						<td>
							<input type="text" id="toreserveddate" name="toreserveddate" value="<%=sReservedToDate%>" size="10" maxlength="10">&nbsp;
							<a href="javascript:void doCalendar('toreserveddate');"><img src="../images/calendar.gif" height="16" width="16" border="0"></a>
							&nbsp;
							<%DrawDateChoices "reserveddate" %>
						</td>
					</tr>
					<tr>
						<td><strong>Reservation Date</strong></td>
						<td>From: </td>
						<td>
							<input type="text" id="fromreservationdate" name="fromreservationdate" value="<%=sReservationFromDate%>" size="10" maxlength="10">&nbsp;
							<a href="javascript:void doCalendar('fromreservationdate');"><img src="../images/calendar.gif" height="16" width="16" border="0"></a>
						</td>
						<td>To: </td>
						<td>
							<input type="text" id="toreservationdate" name="toreservationdate" value="<%=sReservationToDate%>" size="10" maxlength="10">&nbsp;
							<a href="javascript:void doCalendar('toreservationdate');"><img src="../images/calendar.gif" height="16" width="16" border="0"></a>
							&nbsp;
							<%DrawDateChoices "reservationdate" %>
						</td>
					</tr>
					<tr>
						<td colspan="5">
							<input type="button" class="button" id="searchbtn" value="Search" />
						</td>
					</tr>
				</table>
			</form>
	  </fieldset>
<%
		ShowReservations iUserId, sReservedFromDate, sReservedToDate, sReservationFromDate, sReservationToDate
 %>				  
	</div>
</div>
<!--#Include file="../admin_footer.asp"-->  
</body>
</html>
<%
'------------------------------------------------------------------------------------------------------------
'  Function and Subroutines
'------------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' void DrawDateChoices sName
'------------------------------------------------------------------------------------------------------------
Sub DrawDateChoices( ByVal sName )

	response.write vbcrlf & "<select onChange=""getDates(this.value, '" & sName & "');"" class=""calendarinput"" name=""" & sName & """>"
	response.write vbcrlf & "<option value=""0"">Or Select Date Range from Dropdown...</option>"
	response.write vbcrlf & "<option value=""16"">Today</option>"
	response.write vbcrlf & "<option value=""17"">Yesterday</option>"
	response.write vbcrlf & "<option value=""18"">Tomorrow</option>"
	response.write vbcrlf & "<option value=""11"">This Week</option>"
	response.write vbcrlf & "<option value=""12"">Last Week</option>"
	response.write vbcrlf & "<option value=""14"">Next Week</option>"
	response.write vbcrlf & "<option value=""1"">This Month</option>"
	response.write vbcrlf & "<option value=""2"">Last Month</option>"
	response.write vbcrlf & "<option value=""13"">Next Month</option>"
	response.write vbcrlf & "<option value=""3"">This Quarter</option>"
	response.write vbcrlf & "<option value=""4"">Last Quarter</option>"
	response.write vbcrlf & "<option value=""15"">Next Quarter</option>"
	response.write vbcrlf & "<option value=""6"">Year to Date</option>"
	response.write vbcrlf & "<option value=""5"">Last Year</option>"
	response.write vbcrlf & "<option value=""7"">All Dates to Date</option>"
	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void ShowReservations iUserId, sReservedFromDate, sReservedToDate, sReservationFromDate, sReservationToDate
'------------------------------------------------------------------------------------------------------------
Sub ShowReservations( ByVal iUserId, ByVal sReservedFromDate, ByVal sReservedToDate, ByVal sReservationFromDate, ByVal sReservationToDate )
	Dim sSql, oRs, sWhere, iRowCount, iReservationTotal, iRefundTotal, iPaidTotal, iBalanceTotal

	iRowCount = 0
	iReservationTotal = CDbl(0.00)
	iRefundTotal = CDbl(0.00)
	iPaidTotal = CDbl(0.00)
	iBalanceTotal = CDbl(0.00)


	sWhere = " AND R.reserveddate BETWEEN '" & sReservedFromDate & "' AND '" & DateAdd("d",1,sReservedToDate) & "' "

	If sReservationFromDate <> "" Then
		sWhere = sWhere & " AND R.reservationid IN (SELECT reservationid FROM egov_rentalreservationdates "
		sWhere = sWhere & "WHERE reservationstarttime BETWEEN '" & sReservationFromDate & "' AND '" & DateAdd("d",1,sReservationToDate) & "' "
		sWhere = sWhere & "AND orgid = " & session("orgid") & ")"
	End If 

	sSql = "SELECT R.reservationid, R.reserveddate, S.reservationstatus, R.isonhold, ISNULL(R.totalamount,0.00) AS totalamount, "
	sSql = sSql & "ISNULL(R.totalpaid,0.00) AS totalpaid, ISNULL(R.totalrefunded,0.00) AS totalrefunded, "
	sSql = sSql & "ISNULL(R.totalrefundfees,0.00) AS totalrefundfees, ISNULL(R.originalrentalid,0) AS originalrentalid, "
	sSql = sSql & "RN.rentalname, L.name AS locationname "
	sSql = sSql & "FROM egov_rentalreservations R, egov_rentalreservationtypes T, egov_rentalreservationstatuses S, "
	sSql = sSql & "egov_rentals RN, egov_class_location L "
	sSql = sSql & "WHERE R.reservationtypeid = T.reservationtypeid AND R.reservationstatusid = S.reservationstatusid "
	sSql = sSql & "AND R.originalrentalid = RN.rentalid AND RN.locationid = L.locationid "
	sSql = sSql & "AND T.isreservation = 1 AND T.reservationtypeselector = 'public' " & sWhere
	sSql = sSql & " AND R.orgid = " & session("orgid") & " AND R.rentaluserid = " & iUserId
	sSql = sSql & " ORDER BY reserveddate DESC"

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 

		response.write vbcrlf & "<div id=""reservationreportshadow"">"
		response.write vbcrlf & "<table border=""0"" cellspacing=""0"" cellpadding=""2"" id=""reservationreport"">"
		response.write vbcrlf & "<tr class=""tablelist"">"
		response.write "<th align=""center"">Rental</th><th>Reserved<br />On</th><th>Earliest<br />Reservation</th><th>Status</th><th>Reservation</th>"
		response.write "<th>Total<br />Charges</th><th>Total<br />Paid</th><th>Refund<br />Amount</th>"
		response.write "</tr>"

		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"" "
			End If 
			response.write ">"

			' Rental Name
			response.write "<td class=""rentalnamecell"">" & oRs("rentalname") & " &ndash; " & oRs("locationname") & "</td>"
			
			' Reserved Date
			response.write "<td align=""center"">" & DateValue(oRs("reserveddate")) & "</td>"

			' Earliest Reservation Date
			response.write "<td align=""center"">" & GetEarliestReservationDate( oRs("reservationid") )
			response.write "</td>"

			' Status
			response.write "<td align=""center"">" & oRs("reservationstatus") & "</td>"

			' Reservation id to edit link
			response.write "<td align=""center""><a href=""../rentals/reservationedit.asp?reservationid=" & oRs("reservationid") & """>" & oRs("reservationid") & "</a></td>"

			' Total Charges
			response.write "<td align=""right"">" & FormatNumber(oRs("totalamount"),2) & "</td>"
			iReservationTotal = iReservationTotal + CDbl(oRs("totalamount"))

			' Total Paid
			response.write "<td align=""right"">" & FormatNumber(oRs("totalpaid"),2) & "</td>"
			iPaidTotal = iPaidTotal + CDbl(oRs("totalpaid"))

			' Refund Amount
			response.write "<td align=""right"">" & FormatNumber(oRs("totalrefunded"),2) & "</td>"
			iRefundTotal = iRefundTotal + CDbl(oRs("totalrefunded"))

			response.write "</tr>"
			oRs.MoveNext
		Loop 
		
		' Totals Row
		response.write vbcrlf & "<tr id=""totalsrow""><td colspan=""5"" align=""right"">Totals:</td>"
		response.write "<td align=""right"">" & FormatNumber(iReservationTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(iPaidTotal,2) & "</td>"
		response.write "<td align=""right"">" & FormatNumber(iRefundTotal,2) & "</td></tr>"

		response.write vbcrlf & "</table>"
		response.write vbcrlf & "</div>"
	Else
		response.write "<p><strong>No Reservations could be found to match your search criteria.</strong></p>"
	End If 

	oRs.Close
	Set oRs = Nothing 
	
End Sub 


'------------------------------------------------------------------------------------------------------------
' string GetEarliestReservationDate( iReservationid )
'------------------------------------------------------------------------------------------------------------
Function GetEarliestReservationDate( ByVal iReservationid )
	Dim sSql, oRs

	sSql = "SELECT MIN(reservationstarttime) AS reservationstarttime "
	sSql = sSql & "FROM egov_rentalreservationdates WHERE reservationid = " & iReservationid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If Not IsNull(oRs("reservationstarttime")) Then 
			GetEarliestReservationDate = DateValue(oRs("reservationstarttime"))
		Else
			GetEarliestReservationDate = "&nbsp;"
		End If 
	Else
		GetEarliestReservationDate = "&nbsp;"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 



%>
