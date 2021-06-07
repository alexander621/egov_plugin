<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitunscheduledinspections.asp
' AUTHOR: Steve Loar
' CREATED: 05/18/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permit inspections that have not been completed
'
' MODIFICATION HISTORY
' 1.0   05/18/2010	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromInspectionDate, sToInspectionDate, sStreetNumber, sStreetName, sPermitNo
Dim iPermitInspectionTypeId, iInspectorUserId, iInspectionStatusId, sDisplayDateRange
Dim iPermitCategoryId, sPermitLocation

Server.ScriptTimeout=200
'Response.Write(Server.ScriptTimeout)
'response.end

sLevel = "../" ' Override of value from common.asp

' USER SECURITY CHECK and feature availability check, all in one call
PageDisplayCheck "unscheduled inspections", sLevel	' In common.asp

' Handle inspection date range. always want some dates to limit the search
'If request("toinspectiondate") <> "" And request("frominspectiondate") <> "" Then
'	sFromInspectionDate = request("frominspectiondate")
'	sToInspectionDate = request("toinspectiondate")
'	sSearch = sSearch & " AND (I.inspecteddate >= '" & request("frominspectiondate") & "' AND I.inspecteddate < '" & DateAdd("d",1,request("toinspectiondate")) & "' ) "
'	sDisplayDateRange = "From: " & request("frominspectiondate") & " &nbsp;To: " & request("toinspectiondate")
'Else
'	' initially set these to yesterday
'	sFromInspectionDate = FormatDateTime(DateAdd("d",-1,Date),2)
'	sToInspectionDate = FormatDateTime(DateAdd("d",-1,Date),2)
'	sDisplayDateRange = ""
'End If 

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

If request("permitinspectiontypeid") <> "" Then
	iPermitInspectionTypeId = CLng(request("permitinspectiontypeid"))
	If iPermitInspectionTypeId > CLng(0) Then
		sSearch = sSearch & " AND I.permitinspectiontypeid = " & iPermitInspectionTypeId
	End If 
End If 

If request("inspectoruserid") <> "" Then 
	iInspectorUserId = CLng(request("inspectoruserid"))
	If iInspectorUserId > CLng(0) Then 
		sSearch = sSearch & " AND I.inspectoruserid = " & iInspectorUserId
	End If 
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

	<script type="text/javascript" src="../scripts/fastinit.js"></script>
	<script language="Javascript" src="../scripts/tablesort2.js"></script>
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
			document.frmPermitSearch.action = 'permitunscheduledinspections.asp';
			document.frmPermitSearch.submit();
		}

		function doExport()
		{
			document.frmPermitSearch.action='permitunscheduledinspectionsexport.asp';
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
				<font size="+1"><strong>Needs Inspection Report</strong></font>
				<br /><br />
				This report shows all inspections that are outstanding for any permits that have not been voided.
			</p>
			<!--END: PAGE TITLE-->

			<!--BEGIN: FILTER SELECTION-->
			<div class="filterselection">
				<fieldset class="filterselection">
				   <legend class="filterselection">Report Options</legend>
					<p>
						<form name="frmPermitSearch" method="post" action="permitunscheduledinspections.asp">
							<input type="hidden" id="isview" name="isview" value="1" />
							<table cellpadding="2" cellspacing="0" border="0">
								<tr>
									<td>Permit Category:</td>
									<td><%	ShowPermitCategoryPicks iPermitCategoryId	' in permitcommonfunctions.asp	%></td>
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
									<td>Inspection Type:</td><td><% ShowPermitInspectionTypes iPermitInspectionTypeId %></td>
								</tr>
								<tr>
									<td>Inspector:</td><td><% ShowPermitInspectors iInspectorUserId %></td>
								</tr>
								<tr>
									<td colspan="2">
										<input class="button ui-button ui-widget ui-corner-all" type="button" value="View Report" onclick="validate();" />&nbsp;&nbsp;
<%										If request("isview") <> "" Then		%>
											<input type="button" class="button ui-button ui-widget ui-corner-all" value="Download to Excel" onClick="doExport();" />
<%										End If		%>
									</td>
								</tr>
							</table>
						</form>
					</p>
				</fieldset>
			</div>
			<!--END: FILTER SELECTION-->


			<!-- Begin: Report Display -->
<%			' if they choose to view the report, then display the inspections
			If request("isview") <> "" Then	
				DisplayInspections sSearch
			Else 
				response.write "<strong>To view the needed inspections, select from the filter options above then click the &quot;View Report&quot; button.</strong>"
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
' void DisplayLargeAddressList sStreetNumber, sStreetName 
'--------------------------------------------------------------------------------------------------
Sub DisplayLargeAddressList( ByVal sStreetNumber, ByVal sStreetName )
	Dim sSql, oRs, sCompareName

	sSQL = "SELECT DISTINCT sortstreetname, ISNULL(residentstreetprefix,'') AS residentstreetprefix, residentstreetname, "
	sSQL = sSQL & " ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection "
	sSQL = sSQL & " FROM egov_residentaddresses "
	sSQL = sSQL & " WHERE orgid = " & session( "orgid" )
	sSQL = sSQL & " AND residentstreetname IS NOT NULL "
	sSQL = sSQL & " ORDER BY sortstreetname "
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

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
' void ShowPermitInspectionTypes iPermitInspectionTypeId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectionTypes( ByVal iPermitInspectionTypeId )
	Dim sSql, oRs

	sSql = "SELECT permitinspectiontypeid, permitinspectiontype FROM egov_permitinspectiontypes "
	sSql = sSql & " WHERE orgid = " & session("orgid") & " ORDER BY permitinspectiontype"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write "<select name=""permitinspectiontypeid"">"
		response.write vbcrlf & "<option value=""0"""
		If CLng(iPermitInspectionTypeId) = CLng(0) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">All Permit Inspection Types</option>"
		Do While NOT oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("permitinspectiontypeid") & """"
			If CLng(iPermitInspectionTypeId) = CLng(oRs("permitinspectiontypeid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("permitinspectiontype") & "</option>"
			oRs.MoveNext
		Loop
		response.write vbcrlf & "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitInspectors iInspectorUserId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitInspectors( ByVal iInspectorUserId )
	Dim sSql, oRs

	sSql = "SELECT userid, firstname, lastname FROM users WHERE orgid = " & session("orgid") & " AND ispermitinspector = 1 "
	sSql = sSQl & " ORDER BY lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""inspectoruserid"">"
		response.write vbcrlf & "<option value=""0"""
		If CLng(iInspectorUserId) = CLng(0) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">All Inspectors</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option "
			If CLng(iInspectorUserId) = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write " value=""" & oRs("userid") & """>" & oRs("firstname") & " " & oRs("lastname") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayInspections sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayInspections( ByVal sSearch )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT I.permitid, I.permitinspectiontype, S.inspectionstatus, I.scheduleddate, I.inspecteddate, "
	sSql = sSql & " ISNULL(I.inspectoruserid,0) AS inspectoruserid, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S, egov_permitaddress A, egov_permits P, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE I.orgid = " & session("orgid") & " AND I.inspecteddate IS NULL AND S.isneedsinspection = 1 "
	sSql = sSql & " AND S.inspectionstatusid = I.inspectionstatusid AND P.isvoided = 0 AND P.permitlocationrequirementid = R.permitlocationrequirementid AND "
	sSql = sSql & " P.permitid = I.permitid AND A.permitid = I.permitid " & sSearch
	sSql = sSql & " ORDER BY P.permitnumberyear, P.permitnumber, P.permitid, I.inspectionorder"
	'response.write sSql & "<br /><br />"
	'response.end
	' AND I.scheduleddate IS NOT NULL -- pulled to get Piqua inspections to show since they skip this 5/14/2010

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<div id=""inspectiontableshadow"">"
		response.write vbcrlf & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" class=""tableadmin sortable"" id=""inspectionreport"">"
		response.write vbcrlf & "<tr><th>Address/Location</th><th>Permit #</th><th>Inspection<br />Type</th><th>Status</th><th>Scheduled<br />Date</th><th>Inspector</th></tr>"

		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr"
			If iRowCount Mod 2 = 0 Then
				response.write " class=""altrow"""
			End If 
			response.write ">"

			response.write "<td nowrap=""nowrap"" class=""firstcol"">"
			Select Case oRs("locationtype")
				Case "address"
					response.write oRs("permitaddress")

				Case "location"
					response.write Replace(oRs("permitlocation"),Chr(10),"<br />")

				Case Else
					response.write "&nbsp;"

			End Select  
			response.write "</td>"

			response.write "<td align=""center"">" & GetPermitNumber( oRs("permitid") ) & "</td>"
			response.write "<td align=""center"">" & oRs("permitinspectiontype") & "</td>"
			response.write "<td align=""center"">" & oRs("inspectionstatus") & "</td>"
			response.write "<td align=""center"">" & oRs("scheduleddate") & "</td>"
			response.write "<td align=""center"">"
			If CLng(oRs("inspectoruserid")) > CLng(0) Then 
				response.write GetAdminName( oRs("inspectoruserid") )
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"
			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop
		
		response.write vbcrlf & "</table></div>"
		response.flush
	Else
		response.write vbcrlf & "<p>No inspections could be found that match your report criteria.</p>"
		response.flush
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>
