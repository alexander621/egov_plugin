<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitapplicantreport.asp
' AUTHOR: Steve Loar
' CREATED: 06/21/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permits issued by applicant - for Somers
'
' MODIFICATION HISTORY
' 1.0   06/21/2010	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromIssuedDate, sToIssuedDate, sStreetNumber, sStreetName, sPermitNo
Dim iPermitTypeId, sApplicant, iPermitStatusId, sDisplayDateRange, iOrderBY, iPermitCategoryId
Dim sPermitLocation

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK
PageDisplayCheck "permit applicant report", sLevel	' In common.asp

' Handle inspection date range. always want some dates to limit the search
If request("toissueddate") <> "" And request("fromissueddate") <> "" Then
	sFromIssuedDate = request("fromissueddate")
	sToIssuedDate = request("toissueddate")
	sSearch = sSearch & " AND (P.issueddate >= '" & request("fromissueddate") & "' AND P.issueddate < '" & DateAdd("d",1,request("toissueddate")) & "' ) "
	sDisplayDateRange = "From: " & request("fromissueddate") & " &nbsp;To: " & request("toissueddate")
Else
	' initially set these to yesterday
	sFromIssuedDate = FormatDateTime(DateAdd("m",-1,Date),2)
	sToIssuedDate = FormatDateTime(Date,2)
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

If request("permitcategoryid") <> "" Then
	iPermitCategoryId = request("permitcategoryid")
	If CLng(iPermitCategoryId) > CLng(0) Then
		sSearch = sSearch & " AND P.permitcategoryid = " & iPermitCategoryId
	End If 
Else 
	iPermitCategoryId = "0"
End If 

If request("permitlocation") <> "" Then
	sPermitLocation = request("permitlocation")
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If 

If request("orderby") <> "" Then 
	iOrderBY = clng(request("orderby"))
Else
	iOrderBY = clng(1)
End If 

'If sSearch <> "" Then 
'	session("sSql") = sSearch
'Else 
'	session("sSql") = ""
'End If 



if request.form("xls") = "true" then
	Response.ContentType = "application/vnd.ms-excel"
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	Response.AddHeader "Content-Disposition", "attachment;filename=Permit_Applicant_Report_" & sDate & ".xls"
	If clng(iOrderBy) = clng(1) Then 
		DisplayIssuedPermits sSearch
	Else
		DisplayIssuedPermitsByType sSearch
	End If 
else %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			if ($("#fromissueddate").val() != '')
			{
				if (! isValidDate($("#fromissueddate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromissueddate").focus();
					return;
				}
			}
			// check the inspection to date
			if ($("#toissueddate").val() != '')
			{
				if (! isValidDate($("#toissueddate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#toissueddate").focus();
					return;
				}
			}
			document.frmPermitSearch.action='permitapplicantreport.asp';
			document.frmPermitSearch.submit();
		}

		function exportReport()
		{
			// check the inspection from date
			if ($("#fromissueddate").val() != '')
			{
				if (! isValidDate($("#fromissueddate").val()))
				{
					alert("The From Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#fromissueddate").focus();
					return;
				}
			}
			// check the inspection to date
			if ($("#toissueddate").val() != '')
			{
				if (! isValidDate($("#toissueddate").val()))
				{
					alert("The To Date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					$("#toissueddate").focus();
					return;
				}
			}
			document.frmPermitSearch.action='permitapplicantreportexport.asp';
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
				<font size="+1"><strong>Permit Applicant Report</strong></font><br />
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitapplicantreport.asp">
							<input type="hidden" name="xls" value="false" />
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Permit Category:</td>
									<td><%	ShowPermitCategoryPicks iPermitCategoryId	' in permitcommonfunctions.asp	%></td>
								</tr>
								<tr>
									<td>Issued Date:</td>
									<td nowrap="nowrap">
										From:
										<input type="text" id="fromissueddate" name="fromissueddate" value="<%=sFromIssuedDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp; To:
										<input type="text" id="toissueddate" name="toissueddate" value="<%=sToIssuedDate%>" size="10" maxlength="10" class="datepicker" />
										&nbsp;
										<%DrawDateChoices "issueddate" %>
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
									<td>Permit Type:</td><td><% ShowPermitTypes iPermitTypeId %></td>
								</tr>
								<tr>
									<td>Permit Status:</td><td><% ShowPermitStatuses iPermitStatusId %></td>
								</tr>
								<tr>
									<td>Applicant:</td><td><input type="text" id="applicant" name="applicant" size="50" maxlength="50" value="<%=sApplicant%>" /></td>
								</tr>
								<tr>
									<td>Order By:</td>
									<td>
										<select name="orderby">
											<option value="1"
<%												If clng(iOrderBY) = clng(1) Then 
													response.write " selected=""selected"" "
												End If 
%>
											>Location</option>
											<option value="2"
<%												If clng(iOrderBY) = clng(2) Then 
													response.write " selected=""selected"" "
												End If 
%>
											>Permit Type</option>
										</select>
									</td>
								</tr>
								<tr>
									<td colspan="2">
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="document.frmPermitSearch.xls.value = 'false';validate();" />&nbsp;&nbsp;
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="Download to Excel" onclick="document.frmPermitSearch.xls.value = 'true';validate();" />&nbsp;&nbsp;
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
				If clng(iOrderBy) = clng(1) Then 
					DisplayIssuedPermits sSearch
				Else
					DisplayIssuedPermitsByType sSearch
				End If 
			Else 
				response.write "<strong>To view the permit applicant report, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
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
<% end if %>


<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayLargeAddressList sStreetNumber, sStreetName
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oAddressList, sCompareName

	sSql = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSql = sSql & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSql = sSql & " FROM egov_residentaddresses "
	sSql = sSql & " WHERE orgid = " & session( "orgid" )
	sSql = sSql & " AND residentstreetname IS NOT NULL "
	sSql = sSql & " ORDER BY sortstreetname "
	
	Set oAddressList = Server.CreateObject("ADODB.Recordset")
	oAddressList.Open sSql, Application("DSN"), 3, 1

	If NOT oAddressList.EOF Then 
		response.write "<input type=""text"" name=""residentstreetnumber"" value=""" & sStreetNumber & """ size=""8"" maxlength=""10"" /> &nbsp; "
		response.write "<select name=""streetname"">" 
		response.write vbcrlf & "<option value=""0000"">Choose street from dropdown...</option>"

		Do While Not oAddressList.EOF
			sCompareName = ""
			If oAddressList("residentstreetprefix") <> "" Then 
				sCompareName = oAddressList("residentstreetprefix") & " " 
			End If 

			sCompareName = sCompareName & oAddressList("residentstreetname")

			If oAddressList("streetsuffix") <> "" Then 
				sCompareName = sCompareName & " "  & oAddressList("streetsuffix")
			End If 

			If oAddressList("streetdirection") <> "" Then 
				sCompareName = sCompareName & " "  & oAddressList("streetdirection")
			End If 

			response.write vbcrlf & "<option value=""" & sCompareName & """"

			If sStreetName = sCompareName Then 
				response.write " selected=""selected"" "
			End If 

			response.write " >"
			response.write sCompareName & "</option>" 
			oAddressList.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oAddressList.Close
	Set oAddressList = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitTypes iPermitTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitTypes( ByVal iPermitTypeId )
	Dim sSql, oRs

	sSql = "SELECT permittypeid, ISNULL(permittype,'') AS permittype, permittypedesc "
	sSql = sSql & " FROM egov_permittypes "
	sSql = sSql & " WHERE orgid = " & session("orgid")
	sSql = sSql & " ORDER BY permittype, permittypedesc, permittypeid"
	'isbuildingpermittype = 1 AND '

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""permittypeid"">"	
		response.write vbcrlf & "<option value=""0"">All Permit Types</option>"
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
' void DisplayIssuedPermits sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayIssuedPermits( ByVal sSearch )
	Dim sSql, oRs, iRowCount, dJobTotal, dPaidTotal

	iRowCount = 0
	dJobTotal = CDbl(0.00)
	dPaidTotal = CDbl(0.00)

	sSql = "SELECT P.permitid, P.issueddate, P.descriptionofwork, P.applieddate, ISNULL(P.jobvalue,0.00) AS jobvalue, "
	sSql = sSql & " ISNULL(P.totalpaid,0.00) AS totalpaid, S.permitstatus, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " T.permittype, T.permittypedesc, A.legaldescription, ISNULL(C.company,'') AS company, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, ISNULL(C.address,'') AS applicantaddress, ISNULL(C.phone,'') AS applicantphone, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitstatuses S, egov_permittypes T, egov_permitcontacts C, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.issueddate IS NOT NULL " & sSearch
	sSql = sSql & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND P.permittypeid = T.permittypeid "
	sSql = sSql & " AND P.permitid = C.permitid AND C.isapplicant = 1 AND P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " ORDER BY A.residentstreetname, A.streetsuffix, A.residentstreetnumber, P.permitnumberyear, P.permitnumberprefix, P.permitnumber, P.permitid"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""issuedpermitreportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""issuedpermitreport"">"
		response.write vbcrlf & "<tr><th>Permit #</th><th>Issued<br />Date</th><th>Permit<br />Type</th>"
		response.write "<th>Description<br />of Work</th><th>Address/Location</th>"
		response.write "<th>Applicant Name</th><th>Applicant Address</th><th>Applicant Phone</th></tr>"
		response.flush

		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"
			response.write "<td align=""center"" nowrap=""nowrap"">" & GetPermitNumber( oRs("permitid") ) & "</td>"
'			response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("permitstatus") & "</td>"
'			
'			response.write "<td align=""center"">" & FormatDateTime(oRs("applieddate"),2) & "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("issueddate"),2) & "</td>"
			response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("permittype") & "</td>"
			response.write "<td>" & oRs("descriptionofwork") & "</td>"
			
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

'			response.write "<td>&nbsp;" & oRs("legaldescription") & "</td>"
'			response.write "<td align=""right"">" & FormatNumber(oRs("jobvalue"),2) & "</td>"
'			dJobTotal = dJobTotal + CDbl(oRs("jobvalue"))
'			response.write "<td align=""right"">" & FormatNumber(oRs("totalpaid"),2) & "</td>"
'			dPaidTotal = dPaidTotal + CDbl(oRs("totalpaid"))
			If oRs("firstname") <> "" Then 
				response.write "<td align=""left"" nowrap=""nowrap"">&nbsp;" & oRs("firstname") & " " & oRs("lastname") & "</td>"
			Else
				response.write "<td align=""left"" nowrap=""nowrap"">&nbsp;" & oRs("company") & "</td>"
			End If 

			response.write "<td nowrap=""nowrap"">&nbsp;" & oRs("applicantaddress") & "</td>"
			response.write "<td nowrap=""nowrap"">&nbsp;" & FormatPhoneNumber( oRs("applicantphone") ) & "</td>"

			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop
		' Totals row
'		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""8"">&nbsp;</td><td align=""right"">" & FormatNumber(dJobTotal,2) & "</td><td align=""right"">" & FormatNumber(dPaidTotal,2) & "</td><td>&nbsp;</td></tr>"
		response.write vbcrlf & "</table></div>"
		response.flush
	Else
		response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"
		response.flush
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayIssuedPermitsByType sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayIssuedPermitsByType( ByVal sSearch )
	Dim sSql, oRs, iRowCount, dJobTotal, dPaidTotal, sOldType, dTypeJobTotal, dTypePaidTotal, bFirstType

	iRowCount = 0
	dJobTotal = CDbl(0.00)
	dPaidTotal = CDbl(0.00)
	sOldType = "NONE"
	dTypeJobTotal = CDbl(0.00)
	dTypePaidTotal = CDbl(0.00)
	bFirstType = True 

	sSql = "SELECT P.permitid, P.issueddate, P.descriptionofwork, P.applieddate, ISNULL(P.jobvalue,0.00) AS jobvalue, ISNULL(P.totalpaid,0.00) AS totalpaid, S.permitstatus, "
	sSql = sSql & " T.permittype, T.permittypedesc, A.legaldescription, ISNULL(C.company,'') AS company, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, ISNULL(C.address,'') AS applicantaddress, ISNULL(C.phone,'') AS applicantphone, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitstatuses S, egov_permittypes T, egov_permitcontacts C "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.issueddate IS NOT NULL " & sSearch
	sSql = sSql & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND P.permittypeid = T.permittypeid "
	sSql = sSql & " AND P.permitid = C.permitid AND C.isapplicant = 1 "
	sSql = sSql & " ORDER BY T.permittype, P.permitnumberyear, P.permitnumberprefix, P.permitnumber, P.permitid"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""issuedpermitreportshadow"">"
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" class=""tableadmin"" id=""issuedpermitreport"">"
'		response.write vbcrlf & "<tr><th>Permit #</th><th>Current<br />Status</th><th>Description<br />Of Work</th><th>Applied<br />Date</th><th>Issued<br />Date</th><th>Permit<br />Type</th>"
'		response.write "<th>Address</th>"
'		response.write "<th>Legal Desc</th><th>Valuation</th><th>Fees<br />Paid</th><th>Applicant</th></tr>"
		response.write vbcrlf & "<tr><th>Permit #</th><th>Issued<br />Date</th><th>Permit<br />Type</th>"
		response.write "<th>Description<br />of Work</th><th>Location</th>"
		response.write "<th>Applicant Name</th><th>Applicant Address</th><th>Applicant Phone</th></tr>"
		response.flush

		Do While Not oRs.EOF
'			If sOldType <> oRs("permittype") Then
'				If Not bFirstType Then 
'					' Print out a sub total row here
'					response.write vbcrlf & "<tr class=""totalrow issuedsubtotalrow""><td colspan=""8"" align=""right"">" & sOldType & "</td><td align=""right"">" & FormatNumber(dTypeJobTotal,2) & "</td><td align=""right"">" & FormatNumber(dTypePaidTotal,2) & "</td><td>&nbsp;</td></tr>"
'				Else
'					bFirstType = False 
'				End If 
'				sOldType = oRs("permittype")
'				dTypeJobTotal = CDbl(0.00)
'				dTypePaidTotal = CDbl(0.00)
'				iRowCount = 0
'			End If 
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"
			response.write "<td align=""center"" nowrap=""nowrap"">" & GetPermitNumber( oRs("permitid") ) & "</td>"
'			response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("permitstatus") & "</td>"
			
'			response.write "<td align=""center"">" & FormatDateTime(oRs("applieddate"),2) & "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("issueddate"),2) & "</td>"
			response.write "<td align=""center"" nowrap=""nowrap"">" & oRs("permittype") & "</td>"
			response.write "<td>" & oRs("descriptionofwork") & "</td>"
			
			response.write "<td nowrap=""nowrap"">&nbsp;" & oRs("permitaddress") & "</td>"

'			response.write "<td>&nbsp;" & oRs("legaldescription") & "</td>"

'			response.write "<td align=""right"">" & FormatNumber(oRs("jobvalue"),2) & "</td>"
'			dJobTotal = dJobTotal + CDbl(oRs("jobvalue"))
'			dTypeJobTotal = dTypeJobTotal + CDbl(oRs("jobvalue"))

'			response.write "<td align=""right"">" & FormatNumber(oRs("totalpaid"),2) & "</td>"
'			dPaidTotal = dPaidTotal + CDbl(oRs("totalpaid"))
'			dTypePaidTotal = dTypePaidTotal + CDbl(oRs("totalpaid"))

			If oRs("firstname") <> "" Then 
				response.write "<td align=""left"" nowrap=""nowrap"">&nbsp;" & oRs("firstname") & " " & oRs("lastname") & "</td>"
			Else
				response.write "<td align=""left"" nowrap=""nowrap"">&nbsp;" & oRs("company") & "</td>"
			End If 

			response.write "<td nowrap=""nowrap"">&nbsp;" & oRs("applicantaddress") & "</td>"
			response.write "<td nowrap=""nowrap"">&nbsp;" & FormatPhoneNumber( oRs("applicantphone") ) & "</td>"

			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop

		' Print out a sub total row here
'		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""8"" align=""right"">" & sOldType & "</td><td align=""right"">" & FormatNumber(dTypeJobTotal,2) & "</td><td align=""right"">" & FormatNumber(dTypePaidTotal,2) & "</td><td>&nbsp;</td></tr>"

		' Totals row
'		response.write vbcrlf & "<tr class=""totalrow""><td colspan=""8"" align=""right"">Total</td><td align=""right"">" & FormatNumber(dJobTotal,2) & "</td><td align=""right"">" & FormatNumber(dPaidTotal,2) & "</td><td>&nbsp;</td></tr>"
		response.write vbcrlf & "</table></div>"
		response.flush
	Else
		response.write vbcrlf & "<p>No permits could be found that match your report criteria.</p>"
		response.flush
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
