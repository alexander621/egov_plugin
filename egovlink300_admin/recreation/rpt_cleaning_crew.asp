<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: rpt_cleaning_crew.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0	01/17/2006	JOHN STULLENBERGER - INITIAL VERSION
' 1.1	10/06/2006	Steve Loar - Security, Header and nav changed
' 1.2	01/14/2010	Steve Loar - Added back the leasee name
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
sLevel = "../" ' Override of value from common.asp

PageDisplayCheck "cleaning crew report", sLevel	' In common.asp

%>

<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../classes/classes.css" />
	<link rel="stylesheet" href="../rentals/rentalsstyles.css" />
	<link rel="stylesheet" href="global_report.css" />
	<link rel="stylesheet" media="print" href="receiptprint.css" />
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css">

	<script src="https://code.jquery.com/jquery-1.9.1.js"></script>
  	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
  	
  	<script src="../scripts/isvaliddate.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>
	<script async src="../classes/tablesort.js"></script>

	<script>
 	<!--

		var showReport = function() {
			if ( datesAreValid() ) {
				document.searchform.submit();
				return true;
			}
			else {
				return false;
			}
		};
		
		var datesAreValid = function() {
			var okToPost = true;
			// check from date
			if ($("#fromDate").val() != "") {
				if (! isValidDate($("#fromDate").val()) ) {
					inlineMsg("fromDate","<strong>Invalid Value: </strong>The 'from date' should be a valid date in the format of MM/DD/YYYY.");
					okToPost = false;
				}
			}
			else {
				inlineMsg("fromDate","<strong>Missing Value: </strong>The 'from date' is required.");
				okToPost = false;
			}
			
			// check to date
			if ($("#toDate").val() != "") {
				if (! isValidDate($("#toDate").val()) ) {
					inlineMsg("toDate","<strong>Invalid Value: </strong>The 'to date' should be a valid date in the format of MM/DD/YYYY.");
					okToPost = false;
				}
			}
			else {
				inlineMsg("toDate","<strong>Missing Value: </strong>The 'to date' is required.");
				okToPost = false;
			}
			
			return okToPost;
			
		};
		
		$(function() {
			$( "#toDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});

		$(function() {
			$( "#fromDate" ).datepicker({
				showOn: "button",
				buttonImage: "../images/calendar.gif",
				buttonImageOnly: true,
				changeMonth: true,
				changeYear: true
			});
		});

  	//-->
	</script>

</head>
<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="idControls" class="noprint">
	<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
</div>


<!--BEGIN PAGE CONTENT-->
<!--<div id=content style="padding:20px;">-->
<div id="content">
<div id="centercontent">

<h3><%=Session("sOrgName")%> Facility Cleaning Crew Report</h3>

<%
' GET SEARCH DATE RANGE
If request.servervariables("REQUEST_METHOD") = "POST" Then
	datStartDate = request("fromDate")
	datEndDate = request("toDate")
Else
	' DEFAULT TODAYS DATE AND SIX DAYS OUT
	datStartDate = Date()
	datEndDate = Dateadd("d",6,datStartDate)
End If

%>

<!--BEGIN: SEARCH OPTIONS-->
	<fieldset id="search">
		<legend><b>Date Range Selection</b></legend>
		<form action="rpt_cleaning_crew.asp" method="post" name="searchform">
			
			<strong>From:</strong>
			<input type=text id="fromDate" name="fromDate" value="<%=datStartDate%>" />
			
			<span class="searchelement"><strong>To:</strong></span>
			<input type=text id="toDate" name="toDate" value="<%=datEndDate%>" />
			
			<span class="searchelement"><input type="button" value="Search" class="button" onclick="showReport()" /></span>
			
		</form>
	</fieldset>

<!--END: SEARCH OPTIONS-->

<p>

<%

' GENERATE REPORT
SubDisplayReport datStartDate, datEnddate

%>

</p>

</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' void SubDisplayReport datStartDate,datEnddate
'--------------------------------------------------------------------------------------------------
Sub SubDisplayReport( ByVal datStartDate, ByVal datEnddate )
	Dim sSql, oRs, bgcolor

	bgcolor = "#eeeeee"

	sSql = "SELECT facilityscheduleid, facilityname, checkindate, beginhour, beginampm, checkintime, checkouttime, endhour, endampm, userfname, userlname "
	sSql = sSql & "FROM rpt_cleaning_Crew WHERE CAST(checkindate as datetime) "
	sSql = sSql & "BETWEEN '" & datStartDate & "' AND '" & datEnddate & "' AND orgid = " & session("orgid")
	sSql = sSql & " ORDER BY facilityname,checkindate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write "<table id=""reservationlist"" cellpadding=""0"" cellspacing=""0"" bprder=""0"" width=""90%"">"
		response.write "<tr><th>Lodge</th><th>Date</th><th>Time</th><th nowrap=""nowrap"">Arrive Time</th><th nowrap=""nowrap"">Depart Time</th><th>POC</th><th>Lessee</th></tr>"

		Do While Not oRs.EOF 
			If bgcolor = "#eeeeee" Then 
				bgcolor = "#ffffff" 
			Else 
				bgcolor = "#eeeeee"
			End If 
			response.write "<tr bgcolor=""" &  bgcolor  & """>"
			response.write "<td nowrap=""nowrap"" align=""left"">&nbsp;" & oRs("facilityname") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("checkindate") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("beginhour") & " " & oRs("beginampm") & " - " & oRs("endhour") & " " & oRs("endampm") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("checkintime") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & oRs("checkouttime") & "</td>"
			response.write "<td nowrap=""nowrap"" align=""center"">" & GetPOC(oRs("facilityscheduleid")) & "</td>"
			response.write "<td align=""center"">" & oRs("userfname") & " " & oRs("userlname") & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop

		response.write "</table>"

	End If

	oRs.Close
	Set oRs = Nothing 

End Sub


'--------------------------------------------------------------------------------------------------
' string GetPOC( iScheduleid )
'--------------------------------------------------------------------------------------------------
Function GetPOC( ByVal iScheduleid )
	Dim sSql, oRs

	sSql = "SELECT dbo.egov_facility_fields.fieldname, dbo.egov_facility_fields.fieldid, dbo.egov_facility_field_values.fieldvalue, "
	sSql = sSql & "dbo.egov_facility_field_values.paymentid "
	sSql = sSql & "FROM dbo.egov_facility_fields LEFT OUTER JOIN dbo.egov_facility_field_values "
	sSql = sSql & "ON dbo.egov_facility_fields.fieldid = dbo.egov_facility_field_values.fieldid "
	sSql = sSql & "WHERE paymentid = " & iScheduleid & " AND  fieldname = 'poc'"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPOC = oRs("fieldvalue")
	Else
		GetPOC = "n/a"
	End If

	oRs.Close
	Set oRs = Nothing 

End Function


%>


