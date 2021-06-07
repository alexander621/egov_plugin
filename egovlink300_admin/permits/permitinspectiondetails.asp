<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectiondetails.asp
' AUTHOR: Steve Loar
' CREATED: 07/31/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This displays the inspection details for printing.
'
' MODIFICATION HISTORY
' 1.0   07/31/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, iPermitId, sSql, oRs, sPermitInspectionType, sInspectionDescription, bIsReinspection
Dim sRequestReceived, sRequestedDate, sScheduledDate, sInspectedDate, sContactPhone, sContact, sInspectedtime
Dim sScheduledtime, sRequestedTime, sRequestedAmPm, iInspectionStatusId, iInspectorUserId, sSchedulingNotes
Dim sScheduledAmPm, sInspectedAmPm, bIsFinal, bPermitIsCompleted, bIsOnHold, sPermitTypeDesc, sPermitNo
Dim sAddress, sPIN, sInspectorName, sDescOfWork, sProposedUse, sPermitContact, sInspectionStatus
Dim sPermitLocation, sLocationType

iPermitInspectionId = CLng(request("permitinspectionid"))

sSql = "SELECT permitid, permitinspectiontype, inspectiondescription, ISNULL(inspectoruserid,0) AS inspectoruserid, isfinal, "
sSql = sSql & " inspectionstatusid, requestreceiveddate, requesteddate, requestedtime, ISNULL(requestedampm,'') AS requestedampm, "
sSql = sSql & " scheduleddate, scheduledtime, ISNULL(scheduledampm, '') AS scheduledampm, inspecteddate, inspectedtime, ISNULL(inspectedampm,'') AS inspectedampm, "
sSql = sSql & " contact, contactphone, isreinspection, schedulingnotes "
sSql = sSql & " FROM egov_permitinspections WHERE permitinspectionid = " & iPermitInspectionId

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	iPermitId = oRs("permitid")
	sPermitInspectionType = oRs("permitinspectiontype")
	sInspectionDescription = oRs("inspectiondescription")
	bIsReinspection = oRs("isreinspection")
	sRequestReceived = oRs("requestreceiveddate")
	sRequestedDate = oRs("requesteddate")
	sScheduledDate = oRs("scheduleddate")
	sInspectedDate = oRs("inspecteddate")
	sContactPhone = oRs("contactphone")
	sContact = oRs("contact")
	sInspectedtime = oRs("inspectedtime")
	sScheduledtime = oRs("scheduledtime")
	sRequestedTime = oRs("requestedtime")
	iInspectionStatusId = oRs("inspectionstatusid")
	iInspectorUserId = oRs("inspectoruserid")
	sSchedulingNotes = oRs("schedulingnotes")
	sRequestedAmPm = oRs("requestedampm")
	sScheduledAmPm = oRs("scheduledampm")
	sInspectedAmPm = oRs("inspectedampm")
	If oRs("isfinal") Then
		bIsFinal = True 
	Else
		bIsFinal = False 
	End If 
End If 

oRs.Close
Set oRs = Nothing 

If sRequestedTime <> "" Then
	sRequestedDate = sRequestedDate & " " & sRequestedTime & " " & sRequestedAmPm
End If 

If sInspectedtime <> "" Then
	sInspectedDate = sInspectedDate & " " & sInspectedtime & " " & sInspectedAmPm
End If 

'bPermitIsCompleted = GetPermitIsCompleted( iPermitId ) '	in permitcommonfunctions.asp

'bIsOnHold = GetPermitIsOnHold( iPermitId ) '	in permitcommonfunctions.asp

sPermitTypeDesc = GetPermitTypeDesc( iPermitId, True ) '	in permitcommonfunctions.asp

sPermitNo = GetPermitNumber( iPermitId ) '	in permitcommonfunctions.asp

sLocationType = GetPermitLocationType( iPermitId ) '	in permitcommonfunctions.asp

sPermitLocation = Replace(GetPermitPermitLocation( iPermitId ),Chr(10),"<br />") '	in permitcommonfunctions.asp

sAddress = GetPermitJobSite( iPermitId ) '	in permitcommonfunctions.asp

sPin = GetPermitJobSitePIN( iPermitId ) '	in permitcommonfunctions.asp

sInspectorName = GetAdminName( iInspectorUserId ) '	in common.asp

sPermitInspectionType = sPermitInspectionType & " - " & sInspectionDescription

sDescOfWork = GetDescriptionOfWork( iPermitId ) '	in permitcommonfunctions.asp

sProposedUse = GetProposedUse( iPermitId ) '	in permitcommonfunctions.asp

sPermitContact = GetPermitContactAndPhone( iPermitId, "isprimarycontractor" ) '	in permitcommonfunctions.asp

sInspectionStatus = GetInspectionStatusById( iInspectionStatusId ) '	in permitcommonfunctions.asp


%>

<html>
<head>
	<title>E-Gov Permit Inspection Details</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="Javascript">
	<!--

		function doClose()
		{
			parent.hideModal(window.frameElement.getAttribute("data-close"));
		}

	//-->
	</script>

</head>

<body>
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		
	<!--BEGIN: PAGE TITLE-->
	<p>
		Office of the Building Inspector<br />
		<span id="permitinspectionbanner">Permit Inspection Details</span><br /><br />
	</p>
	<!--END: PAGE TITLE-->

	<table cellpadding="4" cellspacing="0" border="0" id="inspectiondetailsreport">
	<tr>
		<td class="inspectiondetailheader" colspan="2">Permit Reference</td>
		<td class="inspectiondetailheader inspectiondetailcentercol" colspan="2">Request</td>
	</tr>
	<tr>
		<td class="inspectiondetaillabel">Permit No</td><td class="inspectiondetails" nowrap="nowrap"><%=sPermitNo%></td>
		<td class="inspectiondetaillabel inspectiondetailcentercol">Inspection Type</td><td class="inspectiondetails"><%=sPermitInspectionType%></td>
	</tr>
	<tr>
		<td class="inspectiondetaillabel">Permit Type</td><td class="inspectiondetails" nowrap="nowrap"><%=sPermitTypeDesc%></td>
		<td class="inspectiondetaillabel inspectiondetailcentercol">Date of Request</td><td class="inspectiondetails"><%=sRequestReceived%></td>
	</tr>
	<tr>
<%		If sLocationType = "address" Then	%>
			<td class="inspectiondetaillabel">Address</td><td class="inspectiondetails" nowrap="nowrap"><%=sAddress%></td>
<%		End If 
		If sLocationType = "location" Then	%>
			<td valign="top" class="inspectiondetaillabel">Location</td><td class="inspectiondetails" nowrap="nowrap"><%=sPermitLocation%></td>
<%		End If		%>
		<td class="inspectiondetaillabel inspectiondetailcentercol">Date Desired</td><td class="inspectiondetails"><%=sRequestedDate%></td>
	</tr>
	<tr>
		<td class="inspectiondetaillabel">PIN</td><td class="inspectiondetails" nowrap="nowrap"><%=sPin%></td>
		<td class="inspectiondetailheader inspectiondetailcentercol" colspan="2">Results</th>
	</tr>
	<tr>
		<td class="inspectiondetaillabel">Description of Work</td><td class="inspectiondetails" nowrap="nowrap"><%=sDescOfWork%></td>
		<td class="inspectiondetaillabel inspectiondetailcentercol">Inspector</td><td class="inspectiondetails"><%=sInspectorName%></td>
	</tr>
	<tr>
		<td class="inspectiondetaillabel">Proposed Use</td><td class="inspectiondetails" nowrap="nowrap"><%=sProposedUse%></td>
		<td class="inspectiondetaillabel inspectiondetailcentercol">Inspected Date</td><td class="inspectiondetails"><%=sInspectedDate%></td>
	</tr>
	<tr>
		<td class="inspectiondetaillabel" valign="top">Primary Contractor</td><td class="inspectiondetails" nowrap="nowrap"><%=sPermitContact%></td>
		<td class="inspectiondetaillabel inspectiondetailcentercol">Result</td><td class="inspectiondetails"><%=sInspectionStatus%></td>
	</tr>
	<tr>
		<td class="inspectiondetaillabel"></td><td class="inspectiondetails">&nbsp;</td>
		<td class="inspectiondetaillabel inspectiondetailcentercol" valign="top">Inspector Notes</td><td class="inspectiondetails"><% ShowInspectorPublicNotes iPermitInspectionId %> </td>
	</tr>
	</table>


	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub ShowInspectorPublicNotes( iPermitReviewId )
'--------------------------------------------------------------------------------------------------
Sub ShowInspectorPublicNotes( iPermitInspectionId )
	Dim sSql, oRs

	sSql = "SELECT entrydate, ISNULL(externalcomment,'') AS externalcomment "
	sSql = sSQl & " FROM egov_permitlog "
	sSql = sSQl & " WHERE externalcomment IS NOT NULL AND isinspectionentry = 1 AND permitinspectionid = " & iPermitInspectionId
	sSql = sSQl & " ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		response.write vbcrlf & FormatDateTime(oRs("entrydate"),2) & " - " & oRs("externalcomment") & "<br />"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 
End Sub 
%>
