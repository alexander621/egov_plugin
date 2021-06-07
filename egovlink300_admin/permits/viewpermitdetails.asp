<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewpermitdetails.asp
' AUTHOR: Steve Loar
' CREATED: 08/05/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  displays permit details
'
' MODIFICATION HISTORY
' 1.0   08/05/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sPermitNo, bIsOnHold, bIsVoided, iPermitStatusId, sLegalDescription, sListedOwner, iPermitAddressId
Dim sApplied, sExpires, sReleased, sApproved, sIssued, sCompleted, sFinishedSqFt, sUnFinishedSqFt, sOtherSqFt
Dim sExaminationHours, sJobValue, sTotalSqFt, sFeeTotal, sNonInvoicedTotal, sInvoicedTotal, sWaivedTotal
Dim sPaidTotal, sDueTotal, sAlertMsg, sAlertSetByUser, dAlertDate, sCounty, sParcelid, sPlansBy, sPrimaryContact
Dim iWorkScopeId, iUseClassId, sProposedUse, sApprovedAs, sResidentialUnits, sOccupants, bHasTempCO, bHasCO
Dim sTempCONotes, sCONotes, sStructureLength, sStructureWidth, sStructureHeight, sZoning, sPlanNumber
Dim sDemolishExistingStructure, sLandFillName, sLandFillCity, sLandFillPhone, iPermitLocationRequirementId
Dim sPermitLocation, bNeedsAddress, bNeedsLocation, bPermitIsInBuildingPermitCategory, bOrgHasVolume

iPermitId = CLng(request("permitid"))

sPermitNo = GetPermitNumber( iPermitId )

bPermitIsInBuildingPermitCategory = PermitIsInBuildingPermitCategory( iPermitId )

sProposedUse = ""
sExistingUse = ""
iWorkClassId = 0
iWorkScopeId = 0
iUseClassId = 0
iConstructionTypeId = 0
iOccupancyTypeId = 0
sDescriptionOfWork = ""
sLegalDescription = ""
iUseTypeId = 0 
sApplied = ""
sReleased = ""
sApproved = ""
sIssued = ""
sCompleted = ""
sExpires = ""
sListedOwner = ""
iMaxContractors = 0
iMaxPriorContacts = 0
iPermitAddressId = 0
sJobValue = FormatNumber(0.00,2,,,0)
sTotalSqFt = FormatNumber(0.00,2,,,0)
sFinishedSqFt = FormatNumber(0.00,2,,,0)
sUnFinishedSqFt = FormatNumber(0.00,2,,,0)
sOtherSqFt = FormatNumber(0.00,2,,,0)
sExaminationHours = FormatNumber(0.00,2,,,0)
sFeeTotal = FormatNumber(0.00,2)
iMaxFees = 0
iMaxReviews = 0
iMaxInspections = 0
iMaxAttachments = 0
iPermitStatusId = 1
bCanPrintPermit = False 
bCanChangeExpirationDate = False 
bHasExpirationDate = True 
bWaiveAllFees = False 
iInvoicePicks = CLng(0)
sPermitnotes = ""
bIsCompleted = False 
sPriorStatus = ""
bCanPlaceHolds = False 
sAlertMsg = ""
sAlertSetByUser = ""
dAlertDate = ""
sPrimaryContact = ""
sStructureLength = ""
sStructureWidth = ""
sStructureHeight = ""
sZoning = ""
sPlanNumber = ""
sDemolishExistingStructure = ""
sLandFillName = ""
sLandFillCity = ""
sLandFillPhone = ""
bNeedsAddress = False 
bNeedsLocation = False 
bOrgHasVolume = False 

bIsOnHold = GetPermitIsOnHold( iPermitId )		' in permitcommonfunctions.asp  
bIsVoided = GetPermitIsVoided( iPermitId )		' in permitcommonfunctions.asp
bHasTempCO = GetPermitPermitTypeFlag( iPermitid, "hastempco" )	' in permitcommonfunctions.asp
bHasCO = GetPermitPermitTypeFlag( iPermitid, "hasco" )	' in permitcommonfunctions.asp
bOrgHasVolume = OrgHasFeature("volume total")

iPermitStatusId = GetPermitStatusId( iPermitId )
If bIsOnHold Then 
	sPermitStatus = "Hold"
Else
	If bIsVoided Then 
		sPermitStatus = "Void"
	Else 
		sPermitStatus = GetPermitStatusByStatusId( iPermitStatusId )
	End If 
End If 

GetPermitDetails iPermitId

sInvoicedTotal = GetInvoicedTotal( iPermitId ) 	' in permitcommonfunctions.asp

sPaidTotal = GetPaidTotal( iPermitId ) 	' in permitcommonfunctions.asp

sWaivedTotal = GetWaivedTotal( iPermitId ) 	' in permitcommonfunctions.asp

sDueTotal = FormatNumber(CDbl(sInvoicedTotal) - ( CDbl(sPaidTotal) + CDbl(sWaivedTotal) ),2)

sNonInvoicedTotal = FormatNumber(CDbl(sFeeTotal) - CDbl(sInvoicedTotal),2)

GetLocationRequirements iPermitLocationRequirementId, bNeedsAddress, bNeedsLocation


%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitdetailsprint.css" media="print" />

	<script type="text/javascript" src="../scripts/layers.js"></script>

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
		<font size="+1"><strong>Permit Details</strong></font><br /><br />
	</p>
	<!--END: PAGE TITLE-->

	<p>
		Permit #: <span class="keyinfo"><%=sPermitNo%> &nbsp; &nbsp; &mdash; &nbsp; <%=GetPermitTypeDesc( iPermitId, False ) %></span>
	</p>
	<p>
		Permit Status: <span class="keyinfo"><%=sPermitStatus%></span>
	</p>
<%	If sAlertMsg <> "" Then %>
		<p>
			<fieldset>
				<legend><span class="keyinfo">Alert</span></legend><br />
				<% response.write "<span id=""permitalertmsg"">" & sAlertMsg & "</span><br />Set by " & sAlertSetByUser & " on " & FormatDateTime(dAlertDate,2)  %>	
			</fieldset>
		</p>
<%	End If		%>

	<p>
		<fieldset>
			<legend><span class="keyinfo">Critical Dates</span></legend><br />
			<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
				<tr><th class="firstcell" align="left">&nbsp;Applied</th><th>Released</th><th>Approved</th><th>Permit<br />Issued</th><th>Completed</th><th>Expires</th></tr>
				<tr>
					<td class="firstcell"><span class="detaildata">&nbsp;<%=sApplied%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sReleased%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sApproved%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sIssued%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sCompleted%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sExpires%></span></td>
				</tr>
				</table>
		</fieldset>
	</p>
	<p>
		<fieldset>
			<legend><span class="keyinfo">Details</span></legend><br />

			<table cellpadding="2" border="0" cellspacing="0" class="viewdetails">

<%			If bNeedsAddress Then	%>
				<tr><td nowrap="nowrap" class="labelcell">Job Site Address:</td><td><span class="keyinfo"><span id="jobaddress"><%=GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, False )%></span></span></td></tr>
				<tr><td nowrap="nowrap" class="labelcell">Listed Owner:</td><td><span class="keyinfo1"><span id="listedowner"><%=sListedOwner%></span></span></td></tr>
				<tr><td nowrap="nowrap" class="labelcell">Legal Description:</td><td><span class="keyinfo1"><span id="legaldescription"><%=sLegalDescription%></span></span></td></tr>
				<tr><td class="labelcell" nowrap="nowrap"><% = GetOrgDisplayWithId( session("orgid"), GetDisplayId("address grouping field"), True ) %>:</td><td><%=sCounty%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap">Parcel Id:</td><td><%=sParcelid%></td></tr>
<%			End If	

			If bNeedsLocation Then	%>
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Location:</td><td><%=Replace(sPermitLocation,Chr(10),"<br />")%></td></tr>
<%			End If	

			If bPermitIsInBuildingPermitCategory Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Use Type:</td><td><%=GetPermitUseType( iUseTypeId ) %></td></tr>
<%			End If 

			If PermitHasDetail( iPermitid, "useclassid" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Use Class:</td><td><%=GetPermitUseClass( iUseClassId ) %></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "descriptionofwork" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Description of Work:</td><td><%=sDescriptionOfWork%></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "workclass" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Work Class:</td><td><%=GetPermitWorkClass( iWorkClassId ) %></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "workscope" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Work Scope:</td><td><%=GetPermitWorkScope( iWorkScopeId ) %></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "constructiontype" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Type of Construction:</td><td><% =GetConstructionType( iConstructionTypeId ) %></td></tr>
				<tr><td class="labelcell" nowrap="nowrap">Occupancy Type:</td><td><% =GetOccupancyType( iOccupancyTypeId ) %></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "existinguse" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Existing Use:</td><td><%=sExistingUse%></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "proposeduse" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Proposed Use:</td><td><%=sProposedUse%></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "approvedas" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Approved As:</td><td><%=sApprovedAs%></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "residentalunits" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">New Residential Units:</td><td><%=sResidentialUnits%></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "occupants" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Occupants:</td><td><%=sOccupants%></td></tr>
<%			End If	

			' Fields For Lansing, IL
			If PermitHasDetail( iPermitid, "structuredimensions" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Structure Dimensions:</td><td>Length: <%=sStructureLength%> &nbsp; Width: <%=sStructureWidth%> &nbsp; Height: <%=sStructureHeight%></td>
<%			End If	
			If PermitHasDetail( iPermitid, "zoning" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Zoning:</td><td><%=sZoning%></td></tr>
<%			End If		
			If PermitHasDetail( iPermitid, "planno" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Plan #:</td><td><%=sPlanNumber%></td></tr>
<%			End If		
			If PermitHasDetail( iPermitid, "demolishexisting" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">&nbsp;</td><td><input type="checkbox" readonly="readonly" name="demolishexistingstructure" <%=sDemolishExistingStructure%> /> Demolish Existing Structure</td></tr>
<%			End If		
			If PermitHasDetail( iPermitid, "landfill" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap">Landfill Name:</td><td><%=sLandFillName%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap">Landfill City:</td><td><%=sLandFillCity%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap">Landfill Phone:</td><td><%=sLandFillPhone%></td></tr>
<%			End If	
			' End of Lansing IL fields

			If PermitHasDetail( iPermitid, "permitnotes" ) Then	%>
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Permit Notes:</td><td><%=sPermitnotes%></td></tr>
<%			End If	
			If PermitHasDetail( iPermitid, "tempconotes" ) Then	
				If bHasTempCO Then		%>
					<tr><td class="labelcell" nowrap="nowrap" valign="top" colspan="2">Temporary CO Stipulations, Conditions, Variances:</td></tr>
					<tr><td class="COcell" nowrap="nowrap" valign="top" colspan="2"><%=sTempCONotes%></td></tr>
<%				Else %>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Temporary CO:</td><td>Does not have TempCO</td></tr>
<%				End If			
			Else 
				If bPermitIsInBuildingPermitCategory Then	%>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Temporary CO:</td><td>Does not have details</td></tr>
<%				End If 
			End If	
			If PermitHasDetail( iPermitid, "conotes" ) Then	
				If bHasCO Then		%>
					<tr><td class="labelcell" nowrap="nowrap" valign="top" colspan="2">Cert of Occupancy Stipulations, Conditions, Variances:</td></tr>
					<tr><td class="COcell" nowrap="nowrap" valign="top" colspan="2"><%=sCONotes%></td></tr>
<%				End If			
			End If		
			
			ShowCustomPermitFields iPermitid 
%>
			</table>
		</fieldset>
	</p>
	<p>
		<fieldset>
			<legend><span class="keyinfo">Contacts</span></legend><br />
			<table cellpadding="2" border="0" cellspacing="0" class="viewdetails">
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Applicant:</td><td><%=GetPermitContactDetails( iPermitId, "isapplicant" )%></td></tr>
				<!-- <tr><td class="labelcell" nowrap="nowrap" valign="top">Primary Contact:</td><td><%'=GetPermitContactDetails( iPermitId, "isprimarycontact" )%></td></tr> -->
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Primary Contact:</td><td><%=sPrimaryContact%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Billing Contact:</td><td><%=GetPermitContactDetails( iPermitId, "isbillingcontact" )%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Primary Contractor:</td><td><%=GetPermitContactDetails( iPermitId, "isprimarycontractor" )%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Architect/Engineer:</td><td><%=GetPermitContactDetails( iPermitId, "isarchitect" )%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Plans By:</td><td><%=GetPermitPlansByContact( iPermitId )%></td></tr>
				<tr><td class="labelcell" nowrap="nowrap" valign="top">Other Contractors:</td><td><%=GetPermitContactDetails( iPermitId, "iscontractor" )%></td></tr>
			</table>
		</fieldset>
	</p>
	<div  class="newpage">&nbsp;</div>
	<p>
		<fieldset>
			<legend><span class="keyinfo">Fees</span></legend><br />
			<p>
				<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
					<tr><th class="firstcell">Finished Sq Ft</th><th>Unfinished Sq Ft</th><th>Total Sq Ft</th>
					<% If bOrgHasVolume Then %>
						<th>Volume</th>
					<% else %>
						<th>Other Sq Ft</th>
					<% End If %>
					<th>Job Value</th><th>Examination<br />Hours</th><th>Fee Total</th></tr>
					<tr>
						<td align="center" class="firstcell"><%=FormatNumber(sFinishedSqFt,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sUnFinishedSqFt,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sTotalSqFt,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sOtherSqFt,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sJobValue,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sExaminationHours,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sFeeTotal,2)%></td>
					</tr>
				</table>
			</p>
			<table cellpadding="2" border="0" cellspacing="0" class="viewdetails">
				<tr><th class="firstcell">Category</th><th>Description</th><th>Method</th><th>Fee Amount</th></tr>
<%				ShowPermitFees iPermitId 		%>	
			</table>
		</fieldset>
	</p>
	<p>
		<fieldset>
			<legend><span class="keyinfo">Invoices &amp; Payments</span></legend><br />
			<p>
				<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
					<caption>Balances</caption>
					<tr><th class="firstcell">Total Fees</th><th>Non-Invoiced Fees</th><th>Invoiced Fees</th><th>Total Waived</th><th>Total Paid</th><th>Total Due</th></tr>
					<tr>
						<td align="center" class="firstcell"><%=FormatNumber(sFeeTotal,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sNonInvoicedTotal,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sInvoicedTotal,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sWaivedTotal,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sPaidTotal,2)%></td>
						<td align="center" class="bordercell"><%=FormatNumber(sDueTotal,2)%></td>
					</tr>
				</table>
			</p>
			<p>
			<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
				<caption>Invoices</caption>
				<tr><th class="firstcell">Invoice #</th><th>Date</th><th>Billed To</th><th>Status</th><th>Invoice Total</th><th>Amount<br />Paid/Waived</th></tr>
				<% ShowInvoices iPermitId %>
			</table>
			</p>
			<p>
			<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
				<caption>Payments</caption>
				<tr><th class="firstcell">Payment #</th><th>Date</th><th>Method</th><th>Amount<br />Paid/Waived</th></tr>
				<% ShowPayments iPermitId %>
			</table>
			</p>
		</fieldset>
	</p>
	<p>
		<fieldset>
			<legend><span class="keyinfo">Reviews</span></legend><br />
			<p>
				<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
				<tr><th class="firstcell">Review</th><th>Status</th><th>Date</th><th>Reviewer</th></tr>
<%				ShowReviewList iPermitId 		%>		
			</table>
		</fieldset>
	</p>
	<div  class="newpage">&nbsp;</div>
	<p>
		<fieldset>
			<legend><span class="keyinfo">Inspections</span></legend><br />
			<p>
				<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
				<tr><th class="firstcell">Inspection</th><th>Reinspection</th><th>Status</th><th>Inspected<br />Date</th><th>Inspector</th></tr>
<%				ShowInspectionList iPermitId 		%>									
			</table>
		</fieldset>
	</p>
	
	<p>
		<fieldset>
			<legend><span class="keyinfo">Attachments</span></legend><br />
			<p>
				<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
				<tr><th class="firstcell">Date Added</th><th>Added By</th><th>File Name</th><th>Description</th></tr>
<%				ShowAttachmentList iPermitId 		%>		
			</table>
		</fieldset>
	</p>
	<p>
		<fieldset>
			<legend><span class="keyinfo">Permit Notes</span></legend><br />
			<p>
				<table cellpadding="4" cellspacing="0" border="0" class="viewdetails">
<%				ShowPermitNotes iPermitId		%>
			</table>
		</fieldset>
	</p>

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

'-------------------------------------------------------------------------------------------------
' void GetPermitDetails iPermitId 
'-------------------------------------------------------------------------------------------------
Sub GetPermitDetails( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitnumberprefix, permitnumberyear, ISNULL(permitnumber,0) AS permitnumber, applieddate, waiveallfees, "
	sSql = sSql & " expirationdate, ISNULL(proposeduse, '') AS proposeduse, ISNULL(existinguse, '') AS existinguse, ISNULL(useclassid, 0) AS useclassid, ISNULL(workscopeid, 0) AS workscopeid, "
	sSql = sSql & " ISNULL(workclassid, 0) AS workclassid, ISNULL(constructiontypeid,0) AS constructiontypeid, ISNULL(permitnotes,'') AS permitnotes, "
	sSql = sSql & " ISNULL(occupancytypeid, 0) AS occupancytypeid, ISNULL(descriptionofwork,'') AS descriptionofwork, "
	sSql = sSql & " ISNULL(usetypeid,0) AS usetypeid, releaseddate, approveddate, issueddate, completeddate, ISNULL(jobvalue,0.00) AS jobvalue, "
	sSql = sSql & " ISNULL(totalsqft,0.00) AS totalsqft, ISNULL(finishedsqft,0.00) AS finishedsqft, ISNULL(unfinishedsqft,0.00) AS unfinishedsqft, "
	sSql = sSql & " ISNULL(othersqft,0.00) AS othersqft, ISNULL(examinationhours,0.00) AS examinationhours, ISNULL(feetotal,0.00) AS feetotal, "
	sSql = sSql & " alertmsg, alertsetbyuserid, alertdate, ISNULL(primarycontact, '') AS primarycontact, residentialunits, approvedas, occupants, "
	sSql = sSql & " ISNULL(tempconotes,'') AS tempconotes, ISNULL(conotes,'') AS conotes, "
	sSql = sSql & " ISNULL(structurelength,'') AS structurelength, ISNULL(structurewidth,'') AS structurewidth, ISNULL(structureheight,'') AS structureheight, "
	sSql = sSql & " ISNULL(zoning,'') AS zoning, ISNULL(plannumber,'') AS plannumber, demolishexistingstructure, "
	sSql = sSql & " ISNULL(landfillname,'') AS landfillname, ISNULL(landfillcity,'') AS landfillcity, ISNULL(landfillphone,'') AS landfillphone, "
	sSql = sSql & " ISNULL(permitlocation,'') AS permitlocation, permitlocationrequirementid "
	sSql = sSql & " FROM egov_permits WHERE permitid = " & iPermitId 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("permitnumber")) > CLng(0) Then 
			sPermitNo = GetPermitNumber( iPermitId )
		Else
			sPermitNo = "None"
		End If 
		sApplied = oRs("applieddate")
		If IsNull(oRs("expirationdate")) Then 
			sExpires = "&nbsp;"
		Else 
			sExpires = FormatDateTime(oRs("expirationdate"), 2)
		End If 
		iPermitLocationRequirementId = oRs("permitlocationrequirementid")
		sPermitLocation = oRs("permitlocation")
		sResidentialUnits = oRs("residentialunits")
		sApprovedAs = oRs("approvedas")
		If Not IsNull(oRs("occupants")) Then 
			sOccupants = oRs("occupants")
		Else 
			sOccupants = "&nbsp;"
		End If 
		sTempCONotes = oRs("tempconotes")
		sCONotes = oRs("conotes")
		sProposedUse = oRs("proposeduse")
		sExistingUse = oRs("existinguse")
		iWorkClassId= oRs("workclassid")
		iUseClassId= oRs("useclassid")
		iWorkScopeId= oRs("workscopeid")
		iConstructionTypeId = oRs("constructiontypeid")
		iOccupancyTypeId = oRs("occupancytypeid")
		sDescriptionOfWork= oRs("descriptionofwork")
		iUseTypeId = oRs("usetypeid")
		If Not IsNull(oRs("releaseddate")) Then 
			sReleased = FormatDateTime(oRs("releaseddate"), 2)
		Else
			sReleased = "&nbsp;"
		End If 
		If Not IsNull(oRs("approveddate")) Then 
			sApproved = FormatDateTime(oRs("approveddate"), 2)
		Else
			sApproved = "&nbsp;"
		End If 
		If Not IsNull(oRs("issueddate")) Then 
			sIssued = FormatDateTime(oRs("issueddate"), 2)
		Else
			sIssued = "&nbsp;"
		End If 
		If Not IsNull(oRs("completeddate")) Then 
			sCompleted = FormatDateTime(oRs("completeddate"), 2)
		Else
			sCompleted = "&nbsp;"
		End If 
		sJobValue = FormatNumber(oRs("jobvalue"),2)
		sTotalSqFt = FormatNumber(oRs("totalsqft"),2,,,0)
		sFinishedSqFt = FormatNumber(oRs("finishedsqft"),2,,,0)
		sUnFinishedSqFt = FormatNumber(oRs("unfinishedsqft"),2,,,0)
		sOtherSqFt = FormatNumber(oRs("othersqft"),2,,,0)
		sExaminationHours = FormatNumber(oRs("examinationhours"),2,,,0)
		sFeeTotal = FormatNumber(oRs("feetotal"),2)
		bWaiveAllFees = oRs("waiveallfees")
		sPermitnotes = oRs("permitnotes")
		If Not IsNull(oRs("alertmsg")) Then 
			sAlertMsg = oRs("alertmsg")
			sAlertSetByUser = GetAdminName( oRs("alertsetbyuserid") )
			dAlertDate = FormatDateTime(oRs("alertdate"), 2)
		End If 
		sTempCONotes = oRs("tempconotes")
		sCONotes = oRs("conotes")
		sPrimaryContact = oRs("primarycontact")
		sStructureLength = oRs("structurelength")
		sStructureWidth = oRs("structurewidth")
		sStructureHeight = oRs("structureheight")
		sZoning = oRs("zoning")
		sPlanNumber = oRs("plannumber")
		If oRs("demolishexistingstructure") Then 
			sDemolishExistingStructure = " checked=""checked"" "
		Else
			sDemolishExistingStructure = ""
		End If 
		sLandFillName = oRs("landfillname")
		sLandFillCity = oRs("landfillcity")
		sLandFillPhone = oRs("landfillphone")
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub  


'--------------------------------------------------------------------------------------------------
' string GetPermitUseType( iUseTypeId )
'--------------------------------------------------------------------------------------------------
Function GetPermitUseType( ByVal iUseTypeId )
	Dim sSql, oRs

	sSql = "SELECT usetype FROM egov_permitusetypes WHERE usetypeid = " & iUseTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitUseType = oRs("usetype")
	Else
		GetPermitUseType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string  GetPermitUseClass( iUseClassId )
'--------------------------------------------------------------------------------------------------
Function GetPermitUseClass( ByVal iUseClassId )
	Dim sSql, oRs

	sSql = "SELECT useclass FROM egov_permituseclasses WHERE useclassid = " & iUseClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitUseClass = oRs("useclass")
	Else
		GetPermitUseClass = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitWorkClass( iWorkClassId ) 
'--------------------------------------------------------------------------------------------------
Function GetPermitWorkClass( ByVal iWorkClassId ) 
	Dim sSql, oRs

	sSql = "SELECT workclass FROM egov_permitworkclasses WHERE workclassid = " & iWorkClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitWorkClass = oRs("workclass")
	Else
		GetPermitWorkClass = ""
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitWorkScope( iWorkScopeId ) 
'--------------------------------------------------------------------------------------------------
Function GetPermitWorkScope( ByVal iWorkScopeId ) 
	Dim sSql, oRs

	sSql = "SELECT workscope FROM egov_permitworkscope WHERE workscopeid = " & iWorkScopeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetPermitWorkScope = oRs("workscope")
	Else
		GetPermitWorkScope = ""
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 


'--------------------------------------------------------------------------------------------------
' string GetConstructionType( iConstructionTypeId )
'--------------------------------------------------------------------------------------------------
Function GetConstructionType( ByVal iConstructionTypeId )
	Dim sSql, oRs

	sSql = "SELECT constructiontype FROM egov_constructiontypes WHERE constructiontypeid = " & iConstructionTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetConstructionType = oRs("constructiontype")
	Else
		GetConstructionType = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' string GetOccupancyType( iOccupancyTypeId )
'--------------------------------------------------------------------------------------------------
Function GetOccupancyType( ByVal iOccupancyTypeId )
	Dim sSql, oRs, sReturn

	sReturn = ""

	sSql = "SELECT ISNULL(usegroupcode,'') AS usegroupcode, occupancytype FROM egov_occupancytypes "
	sSql = sSql & " WHERE occupancytypeid = " & iOccupancyTypeId
	'response.write sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("usegroupcode") <> "" Then
			sReturn = oRs("usegroupcode") & " "
		End If 
		sReturn = sReturn & oRs("occupancytype")
	Else
		sReturn = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

	GetOccupancyType = sReturn

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPermitContactDetails( iPermitId, sContactType )
'--------------------------------------------------------------------------------------------------
Function GetPermitContactDetails( ByVal iPermitId, ByVal sContactType )
	Dim sSql, oRs, sDetails

	sDetails = ""
	sSql = " SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(company,'') AS company, ISNULL(address,'') AS address, ISNULL(city,'') AS city, "
	sSql = sSql & " ISNULL(state,'') AS state, ISNULL(zip,'') AS zip, ISNULL(phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts WHERE " & sContactType & " = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY company, lastname, firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		If sDetails <> "" Then 
			sDetails = sDetails & "<br /><br />"
		End If 
		If oRs("firstname") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If oRs("company") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("company") & "</strong><br />" 
		End If 
		If Trim(oRs("address")) <> "" Then 
			sDetails = sDetails & oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			sDetails = sDetails & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			sDetails = sDetails & FormatPhoneNumber( oRs("phone") ) 
		End If 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

	GetPermitContactDetails = sDetails

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPermitFees iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFees( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT F.permitfeeid, F.isrequired, F.includefee, f.isfixturetypefee, ISNULL(F.permitfeeprefix,'') AS permitfeeprefix, F.permitfee, F.isvaluationtypefee, F.isconstructiontypefee, "
	sSql = sSql & " F.feeamount, ISNULL(F.paymentid,0) AS paymentid, M.permitfeemethod, M.isflatfee, M.ismanual, M.isfixture, F.isupfrontfee, M.ishourly "
	sSql = sSql & " FROM egov_permitfees F, egov_permitfeemethods M "
	sSql = sSql & " WHERE F.permitfeemethodid = M.permitfeemethodid AND F.permitid =" & iPermitId
	sSql = sSql & " ORDER BY F.displayorder, F.permitfeeid"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"   
		response.write "<td nowrap=""nowrap"" align=""center"" class=""firstcell"">"  ' Category cell
		If oRs("permitfeeprefix") = "" Then
			response.write "&nbsp;"
		Else 
			response.write oRs("permitfeeprefix")
		End If 
		response.write "</td>"
		response.write "<td class=""bordercell"">"  ' Description cell
		response.write oRs("permitfee")
		response.write "</td>"
		response.write "<td align=""center"" class=""bordercell"">"  ' Method cell
		response.write oRs("permitfeemethod")
		response.write "</td>"
		response.write "<td align=""center"" class=""bordercell"">"  ' Fee amount cell
		response.write FormatNumber(oRs("feeamount"),2)
		response.write "</td>"
		response.write "</tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInvoices iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowInvoices( ByVal iPermitId )
	Dim sSql, oRs, iRecCount

	sSql = "SELECT I.invoiceid, I.invoicedate, I.totalamount, ISNULL(I.paymentid,0) AS paymentid, I.permitcontactid, "
	sSql = sSql & " S.invoicestatus, I.allfeeswaived, S.isvoid FROM egov_permitinvoices I, egov_invoicestatuses S "
	sSql = sSql & " WHERE I.invoicestatusid = S.invoicestatusid AND I.isvoided = 0 AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			response.write "<td align=""center"" class=""firstcell"">" & oRs("invoiceid") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & FormatDateTime(oRs("invoicedate"),2) & "</td>"
			
			response.write "<td align=""center"" class=""bordercell"">"
			response.write GetInvoiceContact( oRs("permitcontactid") )
			response.write "</td>"

			response.write "<td align=""center"" class=""bordercell"">" & oRs("invoicestatus") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & FormatNumber(oRs("totalamount"),2) & "</td>"
			response.write "<td align=""center"" class=""bordercell"">"
			If oRs("allfeeswaived") Then 
				response.write FormatNumber(oRs("totalamount"),2)
			Else
				response.write FormatNumber(GetInvoicePaymentTotal( CLng(oRs("invoiceid")) ),2)   ' in permitcommonfunctions.asp
			End If 
			response.write "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowReviewList iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowReviewList( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT R.permitreviewid, R.permitreviewtype, R.isrequired, R.isincluded, S.reviewstatus, "
	sSql = sSql & " ISNULL(R.revieweruserid,0) AS revieweruserid, R.reviewed "
	sSql = sSql & " FROM egov_permitreviews R, egov_reviewstatuses S "
	sSql = sSql & " WHERE R.reviewstatusid = S.reviewstatusid AND R.permitid = " & iPermitId
	sSql = sSql & " ORDER BY R.revieworder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			' Review type
			response.write "<td class=""firstcell"">" & oRs("permitreviewtype") & "</td>"
			' Status
			response.write "<td align=""center"" class=""bordercell"">" & oRs("reviewstatus") & "</td>"
			' Reviewed
			response.write "<td align=""center"" class=""bordercell"">"
			If IsNull(oRs("reviewed")) Then
				response.write "&nbsp;"
			Else
				response.write FormatDateTime(oRs("reviewed"),2)
			End If 
			response.write "</td>"

			' Reviewer
			response.write "<td align=""center"" class=""bordercell"">"
			If CLng(oRs("revieweruserid")) > CLng(0) Then 
				response.write GetPermitReviewerName( CLng(oRs("revieweruserid")) )
			Else
				response.write "Unassigned"
			End If 
			response.write "</td>"

			response.write "</tr>"

			ShowReviewNotes oRs("permitreviewid") 

			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowInspectionList( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ShowInspectionList( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT I.permitinspectionid, I.permitinspectiontype, I.inspectiondescription, I.isrequired, S.inspectionstatus, "
	sSql = sSql & " I.inspecteddate, I.isreinspection, ISNULL(I.inspectoruserid,0) AS inspectoruserid, isfinal "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S "
	sSql = sSql & " WHERE I.inspectionstatusid = S.inspectionstatusid AND I.permitid = " & iPermitId
	sSql = sSql & " ORDER BY I.inspectionorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
		
			' Inspection
			response.write "<td class=""firstcell"">" & oRs("permitinspectiontype") & " &mdash; " & oRs("inspectiondescription") & "</td>"

			' Reinspection
			response.write "<td align=""center"" class=""bordercell"">"
			If oRs("isreinspection") Then
				response.write "Reinspection"
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"

			' Status
			response.write "<td align=""center"" class=""bordercell"">" & oRs("inspectionstatus") & "</td>"

			' Date
			response.write "<td align=""center"" class=""bordercell"">" 
			If IsNull(oRs("inspecteddate")) Then
				response.write "&nbsp;"
			Else 
				response.write FormatDateTime(oRs("inspecteddate"),2) 
			End If 
			response.write "</td>"

			' Inspector
			response.write "<td align=""center"" class=""bordercell"">"
			If CLng(oRs("inspectoruserid")) > CLng(0) Then 
				response.write GetAdminName( CLng(oRs("inspectoruserid")) )
			Else
				response.write "Unassigned"
			End If 
			response.write "</td>"

			response.write "</tr>"

			ShowInspectionNotes oRs("permitinspectionid")

			oRs.MoveNext
		Loop
	End If 
	
	oRs.Close
	Set oRs = Nothing 

	ShowInspectionList = iRecCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowAttachmentList iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowAttachmentList( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT permitattachmentid, attachmentname, ISNULL(description,'') AS description, "
	sSql = sSql & " ISNULL(adminuserid,0) AS adminuserid, dateadded "
	sSql = sSql & " FROM egov_permitattachments WHERE permitid = " & iPermitId
	sSql = sSql & " ORDER BY 1 DESC"
	
	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr>"
			response.write "<td align=""center"" class=""firstcell"">" & oRs("dateadded") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & GetAdminName( oRs("adminuserid") ) & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & oRs("attachmentname") & "</td>"
			response.write "<td align=""center"" class=""bordercell"">" & oRs("description") & "</td>"
			response.write "</tr>"
			oRs.MoveNext
		Loop
	Else 
		response.write vbcrlf & "<tr><td colspan=""4"">No Attachments</td></tr>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPermitNotes iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPermitNotes( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, "
	sSql = sSql & " ISNULL(externalcomment,'') AS externalcomment, S.permitstatus, ISNULL(L.adminuserid,0) AS adminuserid, "
	sSql = sSql & " ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_permitstatuses S "
	sSql = sSql & " WHERE isactivityentry = 1 AND S.permitstatusid = L.permitstatusid AND permitid = " & iPermitId
	sSql = sSql & " ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF 
		response.write vbcrlf & "<tr>"
		response.write "<td class=""onlycell"">"
		If CLng(oRs("adminuserid")) > CLng(0) then
			response.write GetAdminName( CLng(oRs("adminuserid")) ) ' In common.asp
		Else
			response.write "System Generated"
		End If 
		response.write " &ndash; " & oRs("permitstatus") & " &ndash; " & oRs("entrydate") & "<br />"
		If oRs("activitycomment") <> "" Then 
			response.write oRs("activitycomment") & "<br />"
		End If 
		If oRs("internalcomment") <> "" Then 
			response.write " &nbsp; <strong>Internal Note:</strong> " & oRs("internalcomment") & "<br />"
		End If 
		If oRs("externalcomment") <> "" Then 
			response.write " &nbsp; <strong>Public Note:</strong> " & oRs("externalcomment")
		End If 
		response.write "</td></tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowReviewNotes iPermitReviewId 
'--------------------------------------------------------------------------------------------------
Sub ShowReviewNotes( ByVal iPermitReviewId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, "
	sSql = sSql & " S.reviewstatus, U.firstname, U.lastname, ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_reviewstatuses S, users U "
	sSql = sSql & " WHERE S.reviewstatusid = L.reviewstatusid AND U.userid = L.adminuserid AND permitreviewid = " & iPermitReviewId
	sSql = sSql & " AND L.isreviewentry = 1 ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<tr><td class=""firstcell"" valign=""top"" align=""right"">Notes:</td><td class=""bordercell"" colspan=""3"">"
		Do While Not oRs.EOF 
			iRowCount = iRowCount + 1
			If CLng(iRowCount) > CLng(1) Then
				response.write "<hr />"
			End If 
			response.write oRs("firstname") & " " & oRs("lastname") & " &ndash; " & oRs("reviewstatus") & " &ndash; " & oRs("entrydate") & "<br />"
			If oRs("activitycomment") <> "" Then 
				response.write " &nbsp; " & oRs("activitycomment") & "<br />"
			End If 
			If oRs("internalcomment") <> "" Then 
				response.write " &nbsp; Internal Note: " & oRs("internalcomment") & "<br />"
			End If 
			If oRs("externalcomment") <> "" Then 
				response.write " &nbsp; Public Note: " & oRs("externalcomment") & "<br />"
			End If 
			oRs.MoveNext
		Loop 
		response.write "</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInspectionNotes iPermitInspectionId 
'--------------------------------------------------------------------------------------------------
Sub ShowInspectionNotes( ByVal iPermitInspectionId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT entrydate, ISNULL(internalcomment,'') AS internalcomment, ISNULL(externalcomment,'') AS externalcomment, "
	sSql = sSql & " S.inspectionstatus, U.firstname, U.lastname, ISNULL(activitycomment,'') AS activitycomment "
	sSql = sSql & " FROM egov_permitlog L, egov_inspectionstatuses S, users U "
	sSql = sSql & " WHERE S.inspectionstatusid = L.inspectionstatusid AND U.userid = L.adminuserid AND permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " AND L.isinspectionentry = 1 ORDER BY permitlogid DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<tr><td class=""firstcell"" valign=""top"" align=""right"">Notes:</td><td class=""bordercell"" colspan=""4"">"
		Do While Not oRs.EOF 
			iRowCount = iRowCount + 1
			If CLng(iRowCount) > CLng(1) Then
				response.write "<hr />"
			End If 
			response.write oRs("firstname") & " " & oRs("lastname") & " &ndash; " & oRs("inspectionstatus") & " &ndash; " & oRs("entrydate") & "<br />"
			If oRs("activitycomment") <> "" Then 
				response.write " &nbsp; " & oRs("activitycomment") & "<br />"
			End If 
			If oRs("internalcomment") <> "" Then 
				response.write " &nbsp; Internal Note: " & oRs("internalcomment") & "<br />"
			End If 
			If oRs("externalcomment") <> "" Then 
				response.write " &nbsp; Public Note: " & oRs("externalcomment") & "<br />"
			End If 
			oRs.MoveNext
		Loop 
		response.write "</td></tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' string GetPermitPlansByContact( iPermitId )
'--------------------------------------------------------------------------------------------------
Function GetPermitPlansByContact( ByVal iPermitId )
	Dim sSql, oRs, sDetails

	sDetails = ""
	sSql = "SELECT ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, "
	sSql = sSql & " ISNULL(C.company,'') AS company, ISNULL(C.address,'') AS address, ISNULL(C.city,'') AS city, "
	sSql = sSql & " ISNULL(C.state,'') AS state, ISNULL(C.zip,'') AS zip, ISNULL(C.phone,'') AS phone " 
	sSql = sSql & " FROM egov_permitcontacts C, egov_permits P "
	sSql = sSql & " WHERE C.permitcontactid = P.plansbycontactid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		If oRs("firstname") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("firstname") & " " & oRs("lastname") & "</strong><br />"
		End If 
		If oRs("company") <> "" Then 
			sDetails = sDetails & "<strong>" & oRs("company") & "</strong><br />" 
		End If 
		If Trim(oRs("address")) <> "" Then 
			sDetails = sDetails & oRs("address") & "<br />" 
		End If 
		If Trim(oRs("city")) <> "" Then 
			sDetails = sDetails & oRs("city") & ", " & oRs("state") & " " & oRs("zip") & "<br />"
		End If 
		If Not IsNull(oRs("phone")) And Trim(oRs("phone")) <> "" Then 
			sDetails = sDetails & FormatPhoneNumber( oRs("phone") ) 
		End If 
	End If  

	oRs.Close
	Set oRs = Nothing 

	GetPermitPlansByContact = sDetails

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowPayments iPermitId 
'--------------------------------------------------------------------------------------------------
Sub ShowPayments( ByVal iPermitId )
	Dim sSql, oRs, dTotal

	dTotal = CDbl(0.00) 

	sSql = "SELECT L.paymentid, J.paymentdate, ISNULL(SUM(L.amount),0.00) AS paymenttotal "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_permitinvoices I, egov_class_payment J "
	sSql = sSql & " WHERE I.isvoided = 0 AND L.invoiceid = I.invoiceid AND L.permitid = " & iPermitId
	sSql = sSql & " AND J.paymentid = L.paymentid AND L.ispaymentaccount = 0 "
	sSql = sSql & " GROUP BY L.paymentid, J.paymentdate"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<tr>"
		response.write "<td align=""center"" valign=""top"" class=""firstcell"">" & oRs("paymentid") & "</td>"
		response.write "<td align=""center"" valign=""top"" class=""bordercell"">" & DateValue(CDate(oRs("paymentdate"))) & "</td>"
		
		response.write "<td valign=""top"" class=""bordercell"">"
		' Show payment types and amount
		ShowInvoicePayments oRs("paymentid")

		response.write "</td>"
		response.write "<td align=""right"" valign=""top"" class=""bordercell"">" & FormatNumber(oRs("paymenttotal"),2) & " &nbsp;</td>"
		response.write "</tr>"
		
		dTotal = dTotal + CDbl(oRs("paymenttotal"))

		oRs.MoveNext
	Loop 
	response.write vbcrlf & "<tr><td colspan=""3""align=""right"" class=""firstcell""><strong>Total Payments</strong>&nbsp;</td>"
	response.write "<td align=""right"" class=""bordercell"">" & FormatNumber(dTotal,2) & " &nbsp;</td>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInvoicePayments iPaymentId 
'--------------------------------------------------------------------------------------------------
Sub ShowInvoicePayments( ByVal iPaymentId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(L.amount,0.00) AS amount, P.paymenttypename, P.requirescheckno "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J, egov_paymenttypes P "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND J.paymentid = " & iPaymentId
	sSql = sSql & " AND L.entrytype = 'debit' AND L.paymenttypeid = P.paymenttypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		'response.write vbcrlf & "<br />"
		response.write "&nbsp;" & oRs("paymenttypename") 
		If oRs("requirescheckno") Then 
			response.write " #: " & GetCheckNo( iPaymentId )
		End If 
		response.write " for " & FormatCurrency(oRs("amount"),2)
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowCustomPermitFields iPermitid 
'--------------------------------------------------------------------------------------------------
Sub ShowCustomPermitFields( ByVal iPermitid )
	Dim sSql, oRs, iCount, sDateValue, sMoneyValue, sIntValue, aChoices, x, sChoice, bCheckFirstRadio
	Dim sSelectedValue, aPicks, bHasChecks

	iCount = clng(0)

	sSql = "SELECT P.customfieldid, F.fieldtypebehavior, P.prompt, P.valuelist, P.fieldsize, "
	sSql = sSql & "ISNULL(P.simpletextvalue,'') AS simpletextvalue, ISNULL(P.largetextvalue,'') AS largetextvalue, "
	sSql = sSql & "P.datevalue, moneyvalue, P.intvalue "
	sSql = sSql & "FROM egov_permitcustomfields P, egov_permitfieldtypes F "
	sSql = sSql & "WHERE P.fieldtypeid = F.fieldtypeid AND P.permitid = " & iPermitid
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		iCount = iCount + 1
		response.write vbcrlf & "<tr><td class=""label"" nowrap=""nowrap"" valign=""top"" colspan=""2"">" & oRs("prompt") & "</td></tr>"
		response.write vbcrlf & "<tr><td class=""label"" nowrap=""nowrap"" valign=""top"">&nbsp;</td><td>"

		Select Case oRs("fieldtypebehavior")
			Case "radio"
				response.write oRs("simpletextvalue")
				
			Case "select"
				response.write oRs("simpletextvalue")

			Case "checkbox"
					response.write Replace(oRs("simpletextvalue"),Chr(10),"<br />")

			Case "textbox"
				response.write oRs("simpletextvalue")

			Case "textarea"
				response.write Replace(oRs("largetextvalue"),Chr(10),"<br />")

			Case "date"
				If IsNull(oRs("datevalue")) Then
					sDateValue = ""
				Else 
					sDateValue = DateValue(oRs("datevalue"))
				End If 
				response.write sDateValue

			Case "money"
				If IsNull(oRs("moneyvalue")) Then
					sMoneyValue = ""
				Else 
					sMoneyValue = FormatNumber(oRs("moneyvalue"),2,,,0)
				End If 
				response.write sMoneyValue

			Case "integer"
				If IsNull(oRs("intvalue")) Then
					sIntValue = ""
				Else 
					sIntValue = oRs("intvalue")
				End If 
				response.write sIntValue

		End Select 

		response.write vbcrlf & "</td></tr>"

		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing 

End Sub  




%>
