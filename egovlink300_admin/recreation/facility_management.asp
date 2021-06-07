<!DOCTYPE html>
<!--#Include file="facility_functions.asp"-->
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
' 1.1   01/18/06   Steve Loar - Code added to display facilities
' 1.2	10/06/06	Steve Loar - Security, Header and nav changed
' 1.3	11/05/07	Steve Loar - Added Send Survey Flag and Ajax code
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim oFacilities, iRowCount, sSQLb

sLevel = "../" ' Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "edit facilities" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>

<html lang="en">
<head>
	<meta charset="UTF-8">

	<title>E-Gov Facility Management</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="facility.css" />

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>

	<script language="Javascript">
	  <!--
		function ConfirmDelete(sFacility, iFacilityId) 
		{
			var msg = "Do you wish to delete " + sFacility + "?"
			if (confirm(msg))
			{
				location.href='facility_delete.asp?ifacilityid='+ iFacilityId;
			}
		}

		function changeSurveyFlag( iFacilityId )
		{
			// Fire off the Send Survey change code without any return handler
			doAjax('setfacilitysendsurvey.asp', 'facilityid=' + iFacilityId, '', 'get', '0');
		}

	  //-->
	 </script>
</head>

<body>
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	
	<p>
	<font size="+1"><strong>Recreation: Facility Management</strong></font><br />
	</p>

	<div id="functionlinks">
		<a href="facility_edit.asp?facilityid=0" id="new_facility"><img src="../images/go.gif" align="absmiddle" border="0">&nbsp;New Facility</a>&nbsp;&nbsp;
	</div>

	<table cellpadding="5" cellspacing="0" border="0" class="tableadmin" id="facilitymgt">
		<tr>
			<th>&nbsp;</th><th>Facility</th><th>Display Template</th>
<%			If OrgHasFeature("facility surveys") Then %>
				<th>Send<br />Surveys</th>
<%			End If %>
			<th>Terms</th><th>Waivers</th><th>Rates & Availability</th><th>Manage Rates</th>
		</tr>
<%
		sSQLb = "SELECT facilityid, facilityname FROM egov_facility WHERE orgid = " & Session("OrgID") & " ORDER BY facilityname"
		Set oFacilities = Server.CreateObject("ADODB.Recordset")
		oFacilities.Open sSQLb, Application("DSN"), 3, 1
		
		If oFacilities.EOF Then
			' something about not having any
			response.write "<tr><td colspan=""6"">There are no facilities.</td></tr>"
		Else
			iRowCount = 0
			Do While Not oFacilities.EOF
				' print out the lines here
				iRowCount = iRowCount + 1
				If iRowCOunt Mod 2 = 0 Then
					response.write "<tr class=""alt_row"">"
				Else
					response.write "<tr>"
				End if
%>
				<td class="action"><a href="facility_edit.asp?facilityid=<%=oFacilities("facilityid")%>">Edit</a>&nbsp;&nbsp;
					<a href="javascript:ConfirmDelete('<%=oFacilities("facilityname")%>',<%=oFacilities("facilityid")%>);">Delete</a></td>
				<td><%=oFacilities("facilityname")%></td>
				<td><%=GetFacilityTemplateName(oFacilities("facilityid"))%></td>
				<%			If OrgHasFeature("facility surveys") Then %>
								<td><input type="checkbox" name="sendsurvey" value="<%=oFacilities("facilityid")%>" <% ShowSendSurveyFlag oFacilities("facilityid") %> onclick="changeSurveyFlag(<%=oFacilities("facilityid")%>);" /></td>
				<%			End If %>
				<td><a href="facility_terms.asp?facilityid=<%=oFacilities("facilityid")%>">Terms</a></td>
				<td><a href="facility_waivers.asp?facilityid=<%=oFacilities("facilityid")%>">Waivers</a></td>
				<!--<td><a href="facility_rates.asp?facilityid=<%=oFacilities("facilityid")%>">Rates</a></td>-->
				<td><a href="facility_availability.asp?facilityid=<%=oFacilities("facilityid")%>">Rates & Availability</a></td>
				<td><a href="facility_rates.asp?facilityid=<%=oFacilities("facilityid")%>">Manage Rates</a></td>
				</tr>
<%
				oFacilities.MoveNext
			Loop 
		End If 

		oFacilities.Close
		Set oFacilities = Nothing 
%>
	</table>

</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>



<%
'--------------------------------------------------------------------------------------------------
' ShowSendSurveyFlag iFacilityId 
'--------------------------------------------------------------------------------------------------
Sub ShowSendSurveyFlag( ByVal iFacilityId )
	Dim sSql, oRs

	sSql = "SELECT sendsurveys FROM egov_facility WHERE facilityid = " & iFacilityId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If oRs("sendsurveys") Then 
		response.write " checked=""checked"" "
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>


