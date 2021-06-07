<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: sewerfeereport.asp
' AUTHOR: Steve Loar
' CREATED: 09/29/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permits issued
'
' MODIFICATION HISTORY
' 1.0   09/29/2008	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromDate, sToDate, iPermitCategoryId
Dim iPermitTypeId, sApplicant

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "sewerfeereport", sLevel	' In common.asp

' Handle inspection date range. always want some dates to limit the search
If request("todate") <> "" And request("fromdate") <> "" Then
	sFromDate = request("fromdate")
	sToDate = request("todate")
	sSearch = sSearch & " AND ((P.issueddate >= '" & request("fromdate") & "' AND P.issueddate < '" & DateAdd("d",1,request("todate")) & "' ) OR ( P.applieddate >= '" & request("fromdate") & "' AND P.applieddate < '" & DateAdd("d",1,request("todate")) & "' ))"
Else
	' initially set these to yesterday
	sFromDate = FormatDateTime(DateAdd("m",-1,Date),2)
	sToDate = FormatDateTime(Date,2)
End If 

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = request("permitcategoryid")
	If CLng(iPermitCategoryId) > CLng(0) Then
		sSearch = sSearch & " AND P.permitcategoryid = " & iPermitCategoryId
	End If 
Else 
	iPermitCategoryId = "0"
End If 

If sSearch <> "" Then 
	session("sSql") = sSearch
Else 
	session("sSql") = ""
End If 

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
		<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
		<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="javascript" src="../scripts/modules.js"></script>
	<script language="Javascript" src="../scripts/getdates.js"></script>
	<script language="JavaScript" src="../scripts/isvaliddate.js"></script>

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
  <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
  <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>

	<script language="Javascript">
	<!--

		function doCalendar( sField ) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmPermitSearch", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function validate()
		{
			// check the inspection from date
			if ($("#fromdate").val() != '')
			{
				if (! isValidDate($("#fromdate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromdate").focus();
					return;
				}
			}
			// check the inspection to date
			if ($("#todate").val() != '')
			{
				if (! isValidDate($("#todate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#todate").focus();
					return;
				}
			}
			document.frmPermitSearch.submit();
		}

  $( function() {
    $( ".datepicker" ).datepicker({
      changeMonth: true,
      showOn: "both",
      buttonText: "<i class=\"fa fa-calendar\"></i>",
      changeYear: true
    });
  } );
	//-->
	</script>

</head>

<body>

	 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<div id="idControls" class="noprint">
		<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	</div>

	<!--BEGIN PAGE CONTENT-->
	<div id="content">
		<div id="centercontent">
		<div class="gutters">

			<!--BEGIN: PAGE TITLE-->
			<p>
				<font size="+1"><strong>Sewer Connection Fee Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="sewerfeereport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Permit Category:</td>
									<td><%	ShowPermitCategoryPicks iPermitCategoryId	' in permitcommonfunctions.asp	%></td>
								</tr>
								<tr>
									<td>Date Range:</td>
									<td nowrap="nowrap">
										From:
										<input type="text" id="fromdate" name="fromdate" value="<%=sFromDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp; To:
										<input type="text" id="todate" name="todate" value="<%=sToDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp;
										<%DrawDateChoices "date" %>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										If request("isview") <> "" Then		%>
											<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="location.href='sewerfeereportexport.asp'" />
<%										End If		%>
									</td>
								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->
		<!--</div>-->

			<!-- Begin: Report Display -->
<%			' if they choose to view the report, then display the inspections
			If request("isview") <> "" Then	
				DisplaySewerConnectionFees sSearch
			Else 
				response.write "<strong>To view the sewer connection fee report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
			End If 
%>
			<!-- END: Report Display -->
		</div>
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
' void DisplaySewerConnectionFees sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplaySewerConnectionFees( ByVal sSearch )
	Dim sSql, oRs, iRowCount, dSewerFeeTotal, dFeesTotal

	iRowCount = 0
	dSewerFeeTotal = CDbl(0.00)
	dFeesTotal = CDbl(0.00)

	sSql = "SELECT P.permitid, P.applieddate, P.issueddate, I.invoiceid, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, ISNULL(C.company,'') AS company, C.contacttype, "
	sSql = sSql & " ISNULL(P.feetotal,0.00) AS feetotal, F.permitfeeprefix, F.permitfee, ISNULL(F.feeamount,0.00) AS feeamount "
	sSql = sSql & " FROM egov_permits P, egov_permitcontacts C, egov_permitfees F, egov_permitinvoiceitems I, egov_permitinvoices II, egov_permitfeereportingtypes R "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.permitid = C.permitid AND C.isapplicant = 1 AND P.permitid = F.permitid "
	sSql = sSql & " AND F.feereportingtypeid = R.feereportingtypeid AND R.issewerconnection = 1 AND I.permitid = P.permitid AND F.permitfeeid = I.permitfeeid "
	sSql = sSql & " AND I.invoiceid = II.invoiceid AND II.permitid = P.permitid AND II.isvoided = 0 " & sSearch
	sSql = sSql & " ORDER BY P.applieddate, I.invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""issuedpermitreportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""issuedpermitreport"">"
		response.write vbcrlf & "<tr><th>Open<br />Date</th><th>Close<br />Date</th><th>Permit #</th><th>Fee Cat</th><th>Description</th><th>Fee</th><th>Total<br />Amount</th><th>Invoice</th><th>Applicant</th></tr>"

		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"

			' Open Date
			response.write "<td align=""center"">" & FormatDateTime(oRs("applieddate"),2) & "</td>"

			' Close date
			response.write "<td align=""center"">" 
			If IsNull(oRs("issueddate")) Then 
				response.write "&nbsp;"
			Else
				response.write FormatDateTime(oRs("issueddate"),2) 
			End If 
			response.write "</td>"

			'Permit Number
			response.write "<td align=""center"" nowrap=""nowrap"">" & GetPermitNumber( oRs("permitid") ) & "</td>"

			' Fee Category
			response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("permitfeeprefix") & "</td>"

			' Fee Desctiption 
			response.write "<td align=""left"">" & oRs("permitfee") & "</td>"

			' Fee Amount
			response.write "<td align=""right"">" & FormatNumber(oRs("feeamount"),2) & "</td>"
			dSewerFeeTotal = dSewerFeeTotal + CDbl(oRs("feeamount"))

			' Total of fees for the permit
			response.write "<td align=""right"">" & FormatNumber(oRs("feetotal"),2) & "</td>"
			dFeesTotal = dFeesTotal + CDbl(oRs("feetotal"))

			' Invoice Number
			response.write "<td align=""center"">" & oRs("invoiceid") & "</td>"

			' Applicant
			response.write "<td align=""left"" nowrap=""nowrap"">"
			If oRs("firstname") <> "" Then 
				response.write oRs("firstname") & " " & oRs("lastname")
			Else
				response.write oRs("company")
			End If 
			response.write "</td>"

			response.write "</tr>"
			oRs.MoveNext 
		Loop
		' Totals row
		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""5"">&nbsp;</td><td align=""right"">" & FormatNumber(dSewerFeeTotal,2) & "</td><td align=""right"">" & FormatNumber(dFeesTotal,2) & "</td><td colspan=""2"">&nbsp;</td></tr>"
		response.write vbcrlf & "</table></div>"
	Else
		response.write vbcrlf & "<p>No fees could be found that match your report selection criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
