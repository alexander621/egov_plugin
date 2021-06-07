<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permittypereport.asp
' AUTHOR: Steve Loar
' CREATED: 11/10/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permits by types. THis is a special report made for Piqua
'
' MODIFICATION HISTORY
' 1.0   11/10/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromAppliedDate, sToAppliedDate, sStreetNumber, sStreetName, sPermitNo
Dim iPermitTypeId, sApplicant, iPermitStatusId, sDisplayDateRange, sPermitLocation

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "permit type report", sLevel	' In common.asp

' Handle inspection date range. always want some dates to limit the search
If request("toapplieddate") <> "" And request("fromapplieddate") <> "" Then
	sFromAppliedDate = request("fromapplieddate")
	sToAppliedDate = request("toapplieddate")
	sSearch = sSearch & " AND (P.applieddate >= '" & request("fromapplieddate") & "' AND P.applieddate < '" & DateAdd("d",1,request("toapplieddate")) & "' ) "
	sDisplayDateRange = "From: " & request("fromapplieddate") & " &nbsp;To: " & request("toapplieddate")
Else
	' initially set these to yesterday
	sFromAppliedDate = FormatDateTime(DateAdd("m",-1,Date),2)
	sToAppliedDate = FormatDateTime(Date,2)
	sDisplayDateRange = ""
End If 

' handle address pick
If request("residentstreetnumber") <> "" Then 
	sStreetNumber = request("residentstreetnumber")
	sSearch = sSearch & "AND A.residentstreetnumber = '" & dbsafe(request("residentstreetnumber")) & "' "
End If 
If request("streetname") <> "" And request("streetname") <> "0000" Then 
	sStreetName = request("streetname")
	sSearch = sSearch & " AND (A.residentstreetname = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbsafe(sStreetName) & "' )"
End If 

' handle the permit number
If request("permitno") <> "" Then 
	sPermitNo = Trim(request("permitno"))
	sSearch = sSearch & BuildPermitNoSearch( sPermitNo )	' in permitcommonfunctions.asp
End If 

If request("permittypeid") <> "" Then
	iPermitTypeId = CLng(request("permittypeid"))
	If iPermitTypeId > CLng(0) Then
		sSearch = sSearch & " AND P.permittypeid = " & iPermitTypeId
	End If 
End If 

If request("permitstatusid") <> "" Then
	iPermitStatusId = CLng(request("permitstatusid"))
	If iPermitStatusId > CLng(0) Then
		Select Case iPermitStatusId
			Case 1
				sSearch = sSearch & " AND (S.isissued = 1 OR S.isissuedback = 1) "
			Case 2
				sSearch = sSearch & " AND S.iscompletedstatus = 1 "
			Case 3
				sSearch = sSearch & " AND P.isonhold = 1 "
			Case 4
				sSearch = sSearch & " AND P.isexpired = 1 "
			Case 5
				sSearch = sSearch & " AND P.isvoided = 1 "
		End Select 
	End If 
End If 

If request("applicant") <> "" Then 
	sApplicant = request("applicant")
	sSearch = sSearch & " AND ( C.company LIKE '%" & dbsafe(sApplicant) & "%' OR C.firstname LIKE '%" & dbsafe(sApplicant) & "%' OR C.lastname LIKE '%" & dbsafe(sApplicant) & "%' ) "
End If 

If request("permitlocation") <> "" Then
	sPermitLocation = request("permitlocation")
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If 


'If sSearch <> "" Then 
'	session("sSql") = sSearch
'Else 
'	session("sSql") = ""
'End If 



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
			// check the applied from date
			if ($("#fromapplieddate").val() != '')
			{
				if (! isValidDate($("#fromapplieddate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromapplieddate").focus();
					return;
				}
			}
			// check the applied to date
			if ($("#toapplieddate").val() != '')
			{
				if (! isValidDate($("#toapplieddate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#toapplieddate").focus();
					return;
				}
			}
			document.frmPermitSearch.action='permittypereport.asp';
			document.frmPermitSearch.submit();
		}

		function exportReport()
		{
			// check the inspection from date
			if ($("#fromapplieddate").val() != '')
			{
				if (! isValidDate($("#fromapplieddate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromapplieddate").focus();
					return;
				}
			}
			// check the inspection to date
			if ($("#toapplieddate").val() != '')
			{
				if (! isValidDate($("#toapplieddate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#toapplieddate").focus();
					return;
				}
			}
			document.frmPermitSearch.action='permittypereportexport.asp';
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
			<p id="pagetitle">
				<span id="printdaterange"><font size="+1"><strong><%=sDisplayDateRange%></strong></font></span>
				<font size="+1"><strong>Permit Type Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permittypereport.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Permit Type:</td><td><% ShowPermitTypes iPermitTypeId %></td>
								</tr>
								<tr>
									<td>Received Date:</td>
									<td nowrap="nowrap">
										From:
										<input type="text" id="fromapplieddate" name="fromapplieddate" value="<%=sFromAppliedDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp; To:
										<input type="text" id="toapplieddate" name="toapplieddate" value="<%=sToAppliedDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp;
										<%DrawDateChoices "applieddate" %>
									</td>
								</tr>
								<tr>
									<td>Address:</td><td><%  DisplayLargeAddressList sStreetNumber, sStreetName %></td>
								</tr>
								<tr>
									<td>Location Like:</td><td><input type="text" name="permitlocation" size="100" maxlength="100" value="<%=sPermitLocation%>" /></td>
								</tr>
								<tr>
									<td>Permit #:</td><td><input type="text" name="permitno" size="20" maxlength="20" value="<%=sPermitNo%>" /></td>
								</tr>
								
								<tr>
									<td>Permit Status:</td><td><% ShowPermitStatuses iPermitStatusId %></td>
								</tr>
								<tr>
									<td>Applicant Like:</td><td><input type="text" id="applicant" name="applicant" size="50" maxlength="50" value="<%=sApplicant%>" /></td>
								</tr>
								<tr>
									<td colspan="2">
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										If request("isview") = "notcreatedyet" Then		%>
											<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="exportReport();" />
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
				DisplayPermitTypeReport sSearch
			Else 
				response.write "<strong>To view the permit type report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
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
' void DisplayLargeAddressList sStreetNumber, sStreetName, 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs, sCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"

		Do While Not oRs.EOF
			sCompareName = ""
			If oRs("residentstreetprefix") <> "" Then 
				sCompareName = oRs("residentstreetprefix") & " " 
			End If 

			sCompareName = sCompareName & oRs("residentstreetname")

			If oRs("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetsuffix")
			End If 

			If oRs("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & oRs("streetdirection")
			End If 

			response.write vbcrlf & "<option value=""" & sCompareName & """"

			If sStreetName = sCompareName Then 
				response.write " selected=""selected"" "
			End If 

			response.write " >"
			response.write sCompareName & "</option>" 
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitTypes iPermitTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitTypes( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permittypeid, ISNULL(permittype,'') AS permittype, permittypedesc "
	sSql = sSql & " FROM egov_permittypes "
	sSql = sSql & " WHERE isbuildingpermittype = 1 AND orgid = "& session("orgid")
	sSql = sSql & " ORDER BY permittype, permittypedesc, permittypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""permittypeid"">"	
		'response.write vbcrlf & "<option value=""0"">All Permit Types</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value="""  & oRs("permittypeid") & """"
			If CLng(iPermitTypeId) = CLng(oRs("permittypeid")) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permittype") & " " & oRs("permittypedesc") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write vbcrlf & "There are No Permit Types to select."
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitStatuses iPermitStatusId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitStatuses( ByVal iPermitStatusId )

	response.write vbcrlf & "<select name=""permitstatusid"">"	
	response.write vbcrlf & "<option value=""0"">Any Permit Status</option>"
	
	response.write vbcrlf & "<option value=""1"""
	If CLng(iPermitStatusId) = CLng(1) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Issued</option>"
	response.write vbcrlf & "<option value=""2"""
	If CLng(iPermitStatusId) = CLng(2) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Completed</option>"
	response.write vbcrlf & "<option value=""3"""
	If CLng(iPermitStatusId) = CLng(3) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">On Hold</option>"
	response.write vbcrlf & "<option value=""4"""
	If CLng(iPermitStatusId) = CLng(4) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Expired</option>"
	response.write vbcrlf & "<option value=""5"""
	If CLng(iPermitStatusId) = CLng(5) Then 
		response.write " selected=""selected"" "
	End If 
	response.write ">Voided</option>"

	response.write vbcrlf & "</select>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayPermitTypeReport sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayPermitTypeReport( ByVal sSearch )
	Dim sSql, oRs, iRowCount, dJobTotal, dPaidTotal

	iRowCount = 0
	dJobTotal = CDbl(0.00)
	dPaidTotal = CDbl(0.00)

	sSql = "SELECT P.permitid, P.applieddate, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, T.permittypeid, "
	sSql = sSql & " ISNULL(C.company,'') AS company, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitstatuses S, egov_permittypes T, egov_permitcontacts C, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & sSearch
	sSql = sSql & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND P.permittypeid = T.permittypeid "
	sSql = sSql & " AND P.permitid = C.permitid AND C.isapplicant = 1 AND P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " ORDER BY P.permitnumberyear, P.permitnumberprefix, P.permitnumber, P.permitid"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""issuedpermitreportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""issuedpermitreport"">"
		response.write vbcrlf & "<tr><th>Permit #</th><th>Address/Location</th><th>Applicant</th><th>Date<br />Received</th>"

		' Display the custom field report titles
		ShowCustomFieldReportTitles oRs("permittypeid")

		response.write "</tr>"

		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"
			response.write "<td align=""center"" nowrap=""nowrap"">" & GetPermitNumber( oRs("permitid") ) & "</td>"
			'response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("permitstatus") & "</td>"

			response.write "<td nowrap=""nowrap"" class=""addresscol"">"
			Select Case oRs("locationtype")
				Case "address"
					response.write oRs("permitaddress")

				Case "location"
					response.write Replace(oRs("permitlocation"),Chr(10),"<br />")

				Case Else
					response.write "&nbsp;"

			End Select  
			response.write "</td>"

			If oRs("firstname") <> "" Then 
				response.write "<td align=""center"" nowrap=""nowrap"">&nbsp;" & oRs("firstname") & " " & oRs("lastname") & "</td>"
			Else
				response.write "<td align=""center"" nowrap=""nowrap"">&nbsp;" & oRs("company") & "</td>"
			End If 

			response.write "<td align=""center"">" & FormatDateTime(oRs("applieddate"),2) & "</td>"

			' Display the selected Custom Fields
			ShowCustomFields oRs("permitid")

			response.write "</tr>"

			oRs.MoveNext 
		Loop

		response.write vbcrlf & "</table></div>"
	Else
		response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowCustomFieldReportTitles iPermitTypeId
'--------------------------------------------------------------------------------------------------
Sub ShowCustomFieldReportTitles( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(C.reporttitle ,'') AS reporttitle "
	sSql = sSql & "FROM egov_permitcustomfieldtypes C,egov_permittypes_to_permitcustomfieldtypes T "
	sSql = sSql & "WHERE T.includeonreport = 1 AND C.customfieldtypeid = T.customfieldtypeid "
	sSql = sSql & "AND T.permittypeid = " & iPermitTypeId
	sSql = sSql & " ORDER BY T.customfieldorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write "<th>" & oRs("reporttitle") & "</th>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowCustomFields iPermitId
'--------------------------------------------------------------------------------------------------
Sub ShowCustomFields( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT F.fieldtypebehavior, ISNULL(C.simpletextvalue,'&nbsp;') AS simpletextvalue, ISNULL(C.largetextvalue,'&nbsp;') AS largetextvalue, "
	sSql = sSql & "C.datevalue, C.moneyvalue, C.intvalue "
	sSql = sSql & "FROM egov_permittypes_to_permitcustomfieldtypes T, egov_permitcustomfields C, egov_permitfieldtypes F, egov_permits P "
	sSql = sSql & "WHERE T.includeonreport = 1 AND T.customfieldtypeid = C.customfieldtypeid AND C.permitid = " & iPermitId
	sSql = sSql & " AND C.fieldtypeid = F.fieldtypeid AND T.permittypeid = P.permittypeid AND P.permitid = C.permitid "
	sSql = sSql & "ORDER BY T.customfieldorder"


	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write "<td align=""center"">" 
		
		Select Case oRs("fieldtypebehavior")
			Case "radio"
				response.write oRs("simpletextvalue")

			Case "select"
				response.write oRs("simpletextvalue")

			Case "checkbox"
				response.write Replace(oRs("simpletextvalue"), Chr(10), "<br />")

			Case "textbox"
				response.write oRs("simpletextvalue")

			Case "textarea"
				response.write Replace(oRs("largetextvalue"), Chr(10), "<br />")

			Case "date"
				If IsNull(oRs("datevalue")) Then
					response.write "&nbsp;"
				Else 
					response.write DateValue(oRs("datevalue"))
				End If 

			Case "money"
				If IsNull(oRs("moneyvalue")) Then
					response.write "&nbsp;"
				Else 
					response.write FormatNumber(oRs("moneyvalue"),2,,,0)
				End If 

			Case "integer"
				If IsNull(oRs("intvalue")) Then
					response.write "&nbsp;"
				Else 
					response.write oRs("intvalue")
				End If 

			Case Else 
				response.write oRs("simpletextvalue")

		End Select
		
		response.write "</td>"

		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub



%>
