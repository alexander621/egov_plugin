<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: cleaningcrewreport.asp
' AUTHOR: Steve Loar
' CREATED: 06/24/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Re-make of the cleaning crew report from Facilities
'
' MODIFICATION HISTORY
' 1.0   06/24/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, toDate, fromDate, today, iOrderBy, sShowBlocked

sLevel = "../" ' Override of value from common.asp

' check if page is online and user has permissions in one call not two
PageDisplayCheck "rentals cleaning crew report", sLevel	' In common.asp

fromDate = Request("fromDate")
toDate = Request("toDate")
today = Date()
sShowBlocked = ""
sBlockFilter = ""

' IF EMPTY DEFAULT TO CURRENT TO DATE
If toDate = "" or IsNull(toDate) Then
	toDate = today 
End If

If fromDate = "" or IsNull(fromDate) Then 
	fromDate = today
End If

sSearch = sSearch & " AND D.reservationstarttime BETWEEN '" & CDate(fromDate) & " 0:00 AM' AND '" & DateAdd("d",1,CDate(toDate)) &" 0:00 AM' "

If request("orderby") = "" Then 
	iOrderBy = CLng(1)
Else
	iOrderBy = CLng(request("orderby"))
End If 

If request("showblocked") = "on" Then
	sShowBlocked = " checked=""checked"""
	sBlockFilter = ""
Else
	sShowBlocked = ""
	sBlockFilter = " AND T.isblock = 0"
End If 


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="Javascript" src="scripts/tablesort.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>


	<script language="Javascript">
	<!--

		function doCalendar( sField ) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			var sSelectedDate = $(sField).value;

			// Set the to date value to the from date value
			if (sField == "toDate")
			{
				sSelectedDate = $("fromDate").value;
				//alert( $("fromDate").value );
			}

			//alert( sField + ": " + sSelectedDate );

			eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmCleaningCrew", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function RefreshResults()
		{
			document.frmCleaningCrew.submit();
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
'		<input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
%>
	</div>

<%
'	<object id="factory" viewastext  style="display:none"
'	  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
'	   codebase="../includes/smsx.cab#Version=6,3,434,12">
'	</object>
%>
	<!--END: THIRD PARTY PRINT CONTROL-->

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Cleaning Crew Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Search Options</legend>
					<p>
						<form name="frmCleaningCrew" method="post" action="cleaningcrewreport.asp">
							<table cellpadding="2" cellspacing="0" border="0">
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
									<td>Order BY:</td>
									<td>
<%										ShowOrderByPicks iOrderBy		%>
									</td>
								</tr>
								<tr>
									<td>&nbsp;</td>
									<td><input type="checkbox" name="showblocked" id="showblocked" <%=sShowBlocked%> /> Include Blocked Times</td>
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


<%				'Pull the data here
				ShowCleaningCrewReport sSearch, iOrderBy, sBlockFilter
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
' void ShowCleaningCrewReport sSearch, iOrderBy
'--------------------------------------------------------------------------------------------------
Sub ShowCleaningCrewReport( ByVal sSearch, ByVal iOrderBy, ByVal sBlockFilter )
	Dim sSql, oRs, iRowCount, sRenter, sClassName, sOrderBy

	iRowCount = 0

	If clng(iOrderBy) = clng(1) Then 
		sOrderBy = "L.name, R.rentalname, D.reservationstarttime"
	Else
		sOrderBy = "D.reservationstarttime, L.name, R.rentalname"
	End If 

	sSql = "SELECT R.rentalname, D.reservationstarttime, D.billingendtime, D.actualstarttime, D.actualendtime, "
	sSql = sSql & "ISNULL(RR.pointofcontact, '') AS pointofcontact, ISNULL(RR.rentaluserid,0) AS rentaluserid, "
	sSql = sSql & "ISNULL(RR.timeid,0) AS timeid, T.reservationtypeselector, L.name AS location, T.isblock "
	sSql = sSql & "FROM egov_rentalreservationdates D, egov_rentals R, egov_rentalreservations RR, "
	sSql = sSql & "egov_rentalreservationstatuses S, egov_rentalreservationtypes T, egov_class_location L "
	sSql = sSql & "WHERE D.orgid = " & session("orgid") & " AND D.rentalid = R.rentalid AND D.reservationid = RR.reservationid "
	sSql = sSql & "AND D.statusid = S.reservationstatusid AND S.isreserved = 1 AND L.locationid = R.locationid "
	sSql = sSql & "AND RR.reservationtypeid = T.reservationtypeid " & sBlockFilter & sSearch 
	sSql = sSql & "ORDER BY " & sOrderBy
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	'response.write vbcrlf & "<div id=""reservationlistshadow"" class=""shadow"">"
	response.write vbcrlf & "<table id=""reservationlist"" cellpadding=""1"" cellspacing=""0"" border=""0"" class=""sortable"">"
	response.write vbcrlf & "<tr><th>Rental</th><th>Reservation Time</th><th>Arrival</th><th>Departure</th><th>POC</th><th>Renter</th></tr>"

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		sRenter = ""
		sClassName = ""

		response.write vbcrlf & "<tr"
		If iRowCount Mod 2 = 0 Then
			response.write " class=""altrow"" "
		End If 
		response.write ">"
		response.write "<td class=""firstcol"">" & oRs("location") & " &ndash; " & oRs("rentalname") & "</td>"

		' reservation times
		response.write "<td>" & DateValue(oRs("reservationstarttime")) & " " 
		response.write GetTimePortion( oRs("reservationstarttime") ) & " &ndash; " & GetTimePortion( oRs("billingendtime") )
		response.write "</td>"

		' arrival time
		response.write "<td align=""center"">" & GetTimePortion( oRs("actualstarttime") ) & "</td>"

		' departure time 
		response.write "<td align=""center"">" & GetTimePortion( oRs("actualendtime") ) & "</td>"

		' Point of Contact
		response.write "<td>"
		If oRs("pointofcontact") = "" Then 
			response.write "&nbsp;"
		Else 
			response.write oRs("pointofcontact") 
		End If 
		response.write "</td>"

		' Renter
		If oRs("isblock") Then
			sRenter = "Blocked"
		Else
			If oRs("reservationtypeselector") = "public" Then
				sRenter = GetCitizenName( oRs("rentaluserid") )
			Else
				If oRs("reservationtypeselector") = "admin" Then
					sRenter = GetAdminName( oRs("rentaluserid") )
				Else 
					If oRs("reservationtypeselector") = "class" Then
						sRenter = GetActivityNo( oRs("timeid") )	' In rentalscommonfunctions.asp 
						sClassName = GetClassName( oRs("timeid") )	' In rentalscommonfunctions.asp 
						If Len(sClassName) > 20 Then
							sClassName = Left(sClassName,17) & "..."
						End If 
						sRenter = sRenter & "<br />" & sClassName	
					End If 
				End If 
			End If 
		End If 
		response.write "<td>" & sRenter & "</td>"

		response.write "</tr>"
		oRs.MoveNext
	Loop 

	response.write vbcrlf & "</table>"
	'response.write "</div>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void ShowOrderByPicks iOrderBy 
'------------------------------------------------------------------------------
Sub ShowOrderByPicks( ByVal iOrderBy )
	
	response.write vbcrlf & "<select name=""orderby"">"

	response.write vbcrlf & "<option value=""1"""
	If CLng(1) = CLng(iOrderBy) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Rental, Reservation Time</option>"

	response.write vbcrlf & "<option value=""2"""
	If CLng(2) = CLng(iOrderBy) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Reservation Time, Rental</option>"

	response.write vbcrlf & "</select>"
	
End Sub 




%>
