<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rentaltotals.asp
' AUTHOR: SteveLoar
' CREATED: 03/19/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Report of counts and total dollars
'
' MODIFICATION HISTORY
' 1.0   03/19/2010	Steve Loar - INITIAL VERSION
' 1.1	11/14/2012	Steve Loar - Adding grouping by location
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iSelectedStartYear, iSelectedEndYear, iSelectedStartMonth, iSelectedEndMonth, dStartDate, dEndDate
Dim iTempMonth, iTempYear, iReservationTypeId, sReservationTypeTitle

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "rental totals rpt", sLevel	' In common.asp

If request("startyear") = "" Then
	iSelectedStartYear = Year(Date())
Else
	iSelectedStartYear = request("startyear")
End If 

If request("endyear") = "" Then
	iSelectedEndYear = Year(Date())
Else
	iSelectedEndYear = request("endyear")
End If 

If request("startmonth") = "" Then
	iSelectedStartMonth = 1
Else
	iSelectedStartMonth = request("startmonth")
End If 

If request("endmonth") = "" Then
	iSelectedEndMonth = Month(Date())
Else
	iSelectedEndMonth = request("endmonth")
End If 

dStartDate = CDate(iSelectedStartMonth & "/1/" & iSelectedStartYear & " 00:00 AM")

dEndDate = CDate(iSelectedEndMonth & "/1/" & iSelectedEndYear & " 00:00 AM")

If dEndDate < dStartDate Then 
	' swap the month and year picks so the earlier picks are the start period
	iTempMonth = iSelectedStartMonth
	iSelectedStartMonth = iSelectedEndMonth
	iSelectedEndMonth = iTempMonth
	iTempYear = iSelectedStartYear
	iSelectedStartYear = iSelectedEndYear
	iSelectedEndYear = iTempYear
End If 

If request("reservationtypeid") <> "" Then
	iReservationTypeId = CLng(request("reservationtypeid"))
	sReservationTypeTitle = GetReservationType( iReservationTypeId )   ' in rentalscommonfunctions.asp
	If sReservationTypeTitle <> "" Then 
		sReservationTypeTitle = sReservationTypeTitle & " Only"
	End If 
Else
	iReservationTypeId = 0
	sReservationTypeTitle = ""
End If 



%>
<html>
<head>
  <title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="reporting.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />
	<link rel="stylesheet" type="text/css" href="receiptprint.css" media="print" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>

	<script language="Javascript">
	<!--

		function viewReport()
		{
			document.frmTotalRpt.submit();
		}

	  //-->
	</script>

</head>

<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<form action="rentaltotals.asp" method="post" name="frmTotalRpt">

	<table border="0" cellpadding="10" cellspacing="0" class="start" width="100%">
		<tr>
			<td><font size="+1"><b>Rentals Total Report</b></font></td>
		</tr>
		<tr>
			<td>
				<div class="filterselection">
				<fieldset id="rentalstotalreport">
					<legend><strong>Select</strong></legend>
				
					<!--BEGIN: FILTERS-->
					<!--BEGIN: DATE FILTERS-->
					<p>
					<table border="0" cellpadding="0" cellspacing="0" id="periodpicks">
						<tr>
							<td class="labelcolumn"><strong>Starting Month:</strong></td>
							<td align="left"> <% ShowMonthPicks iSelectedStartMonth, "startmonth" %> &nbsp;
								<% ShowFloatingYearPicks iSelectedStartYear, "startyear" %>
							</td>
						</tr>
						<tr>
							<td class="labelcolumn"><strong>Ending Month:</strong></td>
							<td align="left"> <% ShowMonthPicks iSelectedEndMonth, "endmonth" %> &nbsp;
								<% ShowFloatingYearPicks iSelectedEndYear, "endyear" %>
							</td>
						</tr>
					</table>
					</p>
					<p><strong>Reservation Type: </strong>
<%
						ShowReservationTypeFilter iReservationTypeId, True 
%>

					<p>
						<input class="button" type="button" value="View Report" onclick="viewReport();" />
					</p>

				</fieldset>
				</div>
				<!--END: FILTERS-->
		    </td>
		</tr>
		<tr>
 			<td valign="top">
	  
				<!--BEGIN: DISPLAY RESULTS-->

<%
				DisplaySummary iSelectedStartMonth, iSelectedStartYear, iSelectedEndMonth, iSelectedEndYear, iReservationTypeId, sReservationTypeTitle
%>
				<!-- END: DISPLAY RESULTS -->
      
			</td>
		 </tr>
	</table>
  </form>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------------------------------------
' void ShowFloatingYearPicks iSelectedYear, sPickName
'------------------------------------------------------------------------------------------------------------
Sub ShowFloatingYearPicks( ByVal iSelectedYear, ByVal sPickName )
	Dim iYear

	response.write vbcrlf & "<select name=""" & sPickName & """>"

	For iYear=Year(Date())-5 To Year(Date())+5
		response.write vbcrlf & "<option"
		If clng(iSelectedYear) = clng(iYear) Then
			response.write " selected=""selected"""
		End If
		response.write " value="""& iYear &""">" & iYear & "</option>"
	Next

	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void ShowMonthPicks iSelectedMonth, sPickName
'------------------------------------------------------------------------------------------------------------
Sub ShowMonthPicks( ByVal iSelectedMonth, ByVal sPickName )
	Dim iMonth

	response.write vbcrlf & "<select name=""" & sPickName & """>"

	For iMonth = 1 To 12
		response.write vbcrlf & "<option"
		If clng(iSelectedMonth) = clng(iMonth) Then
			response.write " selected=""selected"""
		End If
		response.write " value="""& iMonth &""">" & MonthName(iMonth) & "</option>"
	Next 

	response.write vbcrlf & "</select>"

End Sub 


'------------------------------------------------------------------------------------------------------------
' void DisplaySummary iSelectedStartMonth, iSelectedStartYear, iSelectedEndMonth, iSelectedEndYear, iReservationTypeId, sReservationTypeTitle
'------------------------------------------------------------------------------------------------------------
Sub DisplaySummary( ByVal iSelectedStartMonth, ByVal iSelectedStartYear, ByVal iSelectedEndMonth, ByVal iSelectedEndYear, ByVal iReservationTypeId, ByVal sReservationTypeTitle )
	Dim sSql, oRs, dStartDate, dEndDate, iOldYear, iOldMonth, dMonthTotal, dGrandTotal, iMonthQty, iGrandQty
	Dim sReservationTypeSelector, bPrintNewMonthCell, sCurrentLocation

	iOldYear = 0
	iOldMonth = 0
	dMonthTotal = CDbl(0.0000)
	dGrandTotal = CDbl(0.0000)
	iMonthQty = CLng(0)
	iGrandQty = CLng(0)
	sCurrentLocation = "None"

	' set the start to the last minute of the prior month
	dStartDate = CDate(iSelectedStartMonth & "/1/" & iSelectedStartYear & " 00:00 AM")
	dStartDate = DateAdd("n", -1, dStartDate)

	' set the end date to the first minute of the following month
	dEndDate = CDate(iSelectedEndMonth & "/1/" & iSelectedEndYear & " 00:00 AM")
	dEndDate = DateAdd("m", 1, dEndDate)

	If CLng(iReservationTypeId) > CLng(0) Then
		sReservationTypeSelector = " AND V.reservationtypeid = " & iReservationTypeId & " "
	Else
		sReservationTypeSelector = ""
	End If 

	' This gets the upper reservation count (no days) and the total for those reservations - There is no duplication here
	sSql = "SELECT L.name AS location, R.rentalname, MONTH(V.reserveddate) AS reservemonth, YEAR(V.reserveddate) AS reserveyear, "
	sSql = sSql & "SUM(V.totalamount) AS totalamount, COUNT(V.reservationid) AS reservationsmade "
	sSql = sSql & "FROM egov_rentals R, egov_rentalreservationstatuses S, egov_rentalreservations V, egov_rentalreservationtypes T, egov_class_location L "
	sSql = sSql & "WHERE V.orgid = " & session("orgid") & " AND V.reserveddate BETWEEN '" & dStartDate & "' AND '" & dEndDate & "' "
	sSql = sSql & "AND V.originalrentalid = R.rentalid AND V.reservationstatusid = S.reservationstatusid AND R.locationid = L.locationid "
	sSql = sSql & "AND S.isreserved = 1 AND V.reservationtypeid = T.reservationtypeid AND T.isreservation = 1 " & sReservationTypeSelector
	sSql = sSql & "GROUP BY L.name, R.rentalname, MONTH(V.reserveddate), YEAR(V.reserveddate) "
	sSql = sSql & "ORDER BY YEAR(V.reserveddate), MONTH(V.reserveddate), L.name, R.rentalname"

	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""rentalscount"" cellpadding=""1"" cellspacing=""0"" border=""0"" class=""sortable"">"
		response.write vbcrlf & "<tr><th colspan=""2"" align=""left"" class=""firstcol"">"
		If sReservationTypeTitle <> "" Then 
			response.write sReservationTypeTitle
		Else 
			response.write "&nbsp;"
		End If 
		response.write "</th><th>Reservations</th><th>Charges</th></tr>"

		Do While Not oRs.EOF
			bPrintNewMonthCell = False 
			If clng(oRs("reserveyear")) <> clng(iOldYear) Or clng(oRs("reservemonth")) <> clng(iOldMonth) Then
				If iOldYear <> clng(0) Then
					' write out the prior month total
					response.write "<tr><td class=""monthtotalrow"">&nbsp;</td><td class=""monthtotalrow""><strong>" & MonthName(iOldMonth) & " " & iOldYear & " Totals</strong></td><td align=""center"" class=""monthtotalrow""><strong>" & iMonthQty & "</strong></td><td align=""right"" class=""monthtotalrow moneycell""><strong>" & FormatNumber(dMonthTotal,2,,,0) & "</strong></td></tr>"
				End If 
				iOldYear = oRs("reserveyear")
				iOldMonth = oRs("reservemonth")
				dMonthTotal = CDbl(0.0000)
				iMonthQty = CLng(0)
				' new month title row
				'response.write "<tr><td class=""firstcol""><strong>" & MonthName(iOldMonth) & " " & iOldYear & "</strong></td><td colspan=""3"">&nbsp;</td></tr>"
				bPrintNewMonthCell = True 
				sCurrentLocation = "None"
			End If 
			dMonthTotal = dMonthTotal + CDbl(oRs("totalamount"))
			dGrandTotal = dGrandTotal + CDbl(oRs("totalamount"))
			iMonthQty = iMonthQty + CLng(oRs("reservationsmade"))
			iGrandQty = iGrandQty + CLng(oRs("reservationsmade"))

			' Write out the detail line for each rental for each month
			'response.write "<tr>"
			If bPrintNewMonthCell Then
				response.write "<tr><td class=""firstcol""><strong>" & MonthName(iOldMonth) & " " & iOldYear & "</strong></td><td colspan=""3"">&nbsp;</td></tr>"
			'Else 
			'	response.write "<tr><td>&nbsp;</td>"
			End If 

			If sCurrentLocation <> oRs("location") Then 
				sCurrentLocation = oRs("location")
				response.write "<tr><td class=""locationcol"" colspan=""4"">" & sCurrentLocation & "</td></tr>"
			End If 

			response.write "<tr><td>&nbsp;</td>"
			response.write "<td>" & Ors("rentalname") & "</td><td align=""center"">" & oRs("reservationsmade") & "</td><td align=""right"" class=""moneycell"">" & FormatNumber(CDbl(oRs("totalamount")),2,,,0) & "</td></tr>"
			oRs.MoveNext
		Loop 

		' Print the last month's total
		response.write "<tr><td class=""finalmonthsum"">&nbsp;</td><td class=""finalmonthsum""><strong>" & MonthName(iOldMonth) & " " & iOldYear & " Totals</strong></td><td align=""center"" class=""finalmonthsum""><strong>" & iMonthQty & "</strong></td><td align=""right"" class=""finalmonthsum moneycell""><strong>" & FormatNumber(dMonthTotal,2,,,0) & "<strong></td></tr>"

		' Print out the total row
		response.write "<tr class=""totalrow""><td>&nbsp;</td><td>Grand Totals</td><td align=""center"">" & iGrandQty & "</td><td align=""right"" class=""moneycell"">" & FormatNumber(dGrandTotal,2,,,0) & "</td></tr>"

		response.write vbcrlf & "</table>"

	Else
		response.write "<p>No information could be found to match your selections. Please change your selections and try again.</p>"

	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
