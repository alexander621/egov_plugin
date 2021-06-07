<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<!-- #include file="permitscommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitdetails.asp
' AUTHOR: Steve Loar
' CREATED: 04/26/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Display the public viewable details of a permit.
'
' MODIFICATION HISTORY
' 1.0   04/26/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sTitle 
Dim iPermitId, sPermitNo, bIsOnHold, bIsVoided, iPermitStatusId, sLegalDescription, sListedOwner, iPermitAddressId
Dim sApplied, sExpires, sReleased, sApproved, sIssued, sCompleted, sFinishedSqFt, sUnFinishedSqFt, sOtherSqFt
Dim sExaminationHours, sJobValue, sTotalSqFt, sFeeTotal, sNonInvoicedTotal, sInvoicedTotal, sWaivedTotal
Dim sPaidTotal, sDueTotal, sAlertMsg, sAlertSetByUser, dAlertDate, sCounty, sParcelid, sPlansBy, sPrimaryContact
Dim iWorkScopeId, iUseClassId, sProposedUse, sApprovedAs, sResidentialUnits, sOccupants, bHasTempCO, bHasCO
Dim sTempCONotes, sCONotes, sStructureLength, sStructureWidth, sStructureHeight, sZoning, sPlanNumber
Dim sDemolishExistingStructure, sLandFillName, sLandFillCity, sLandFillPhone, bIsFound, bIsExpired

If iorgid = 7 Then
	sTitle = sOrgName
Else
	sTitle = "E-Gov Services " & sOrgName
End If

If IsNumeric(request("p")) Then 
	iPermitId = CLng(request("p"))
Else
	response.redirect "permitsearch.asp"
End If 

sPermitNo = GetPermitNumber( iPermitId )

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

bIsOnHold = GetPermitIsOnHold( iPermitId )		' in permitscommonfunctions.asp  
bIsVoided = GetPermitIsVoided( iPermitId )		' in permitscommonfunctions.asp
bIsExpired = GetPermitIsExpired( iPermitId )	' in permitscommonfunctions.asp
bHasTempCO = GetPermitPermitTypeFlag( iPermitid, "hastempco" )	' in permitscommonfunctions.asp
bHasCO = GetPermitPermitTypeFlag( iPermitid, "hasco" )	' in permitscommonfunctions.asp

iPermitStatusId = GetPermitStatusId( iPermitId )
If bIsOnHold Then 
	sPermitStatus = "On Hold"
Else
	If bIsExpired Then
		sPermitStatus = "Expired"
	Else 
		If bIsVoided Then 
			sPermitStatus = "Voided"
		Else 
			sPermitStatus = GetPermitStatusByStatusId( iPermitStatusId )
		End If 
	End If 
End If 

bIsFound = GetPermitDetails( iPermitId )

If bIsFound Then 
	sInvoicedTotal = GetInvoicedTotal( iPermitId ) 	' in permitscommonfunctions.asp

	sPaidTotal = GetPaidTotal( iPermitId ) 	' in permitscommonfunctions.asp

	sWaivedTotal = GetWaivedTotal( iPermitId ) 	' in permitscommonfunctions.asp

	sDueTotal = FormatNumber(CDbl(sInvoicedTotal) - ( CDbl(sPaidTotal) + CDbl(sWaivedTotal) ),2)

	sNonInvoicedTotal = FormatNumber(CDbl(sFeeTotal) - CDbl(sInvoicedTotal),2)
End If 

%>

<html>
<head>

	<title><%=sTitle%></title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permitsstyles.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />
	<link rel="stylesheet" type="text/css" href="permitprintstyles.css" media="print" />

	<script language="JavaScript" src="../prototype/prototype-1.6.0.2.js"></script>
	<script language="JavaScript" src="../scriptaculous/src/scriptaculous.js"></script>

	<script language="Javascript">
	<!--


	//-->
	</script>

</head>

<!--#Include file="../include_top.asp"-->

<!--BEGIN PAGE CONTENT-->

<%	RegisteredUserDisplay( "../" ) %>
<br /><br />

<div id="content">
	<div id="centercontent">
		<div id="topleftbuttons">
			<input type="button" class="button" onclick="javascript:history.go(-1);" value="<< Back" /> &nbsp;&nbsp;
<%	If bIsFound Then		%>
			<input type="button" class="button" onclick="javascript:window.print();" value="Print" />
		</div>



		<!--BEGIN: PAGE TITLE-->
		<p id="titleline">
			<font size="+1"><strong>Permit Details</strong></font><br /><br />
		</p>
		<!--END: PAGE TITLE-->
		<p>
			Permit #: <span class="keyinfo"><%=sPermitNo%> &nbsp; &nbsp; &mdash; &nbsp; <%=GetPermitTypeDesc( iPermitId, False ) %></span>
		</p>
		<p>
			Permit Status: <span class="keyinfo"><%=sPermitStatus%></span>
		</p>
		<p>
			<span class="keyinfo">Critical Dates</span><br />
			<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
				<tr><th class="firstcell" align="left">&nbsp;Applied</th><th>Released</th><th>Approved</th><th>Permit<br />Issued</th><th>Completed</th><th>Expires</th></tr>
				<tr>
					<td class="firstcell"><span class="detaildata">&nbsp;<%=DateValue(sApplied)%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sReleased%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sApproved%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sIssued%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sCompleted%></span></td>
					<td align="center" class="bordercell"><span class="detaildata"><%=sExpires%></span></td>
				</tr>
			</table>
		</p>
		<p>
				<span class="keyinfo">Details</span><br />
				<table cellpadding="2" border="0" cellspacing="0" class="viewdetails">
					<tr><td nowrap="nowrap" class="labelcell">Job Site Address:</td><td><span class="keyinfo"><span id="jobaddress"><%=GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, False )%></span></span></td></tr>
					<tr><td nowrap="nowrap" class="labelcell">Listed Owner:</td><td><span class="keyinfo1"><span id="listedowner"><%=sListedOwner%></span></span></td></tr>
					<tr><td nowrap="nowrap" class="labelcell">Legal Description:</td><td><span class="keyinfo1"><span id="legaldescription"><%=sLegalDescription%></span></span></td></tr>

					<tr><td class="labelcell" nowrap="nowrap"><% = GetOrgDisplayWithId( iOrgid, GetDisplayId("address grouping field"), True ) %>:</td><td><%=sCounty%></td></tr>
					
					<tr><td class="labelcell" nowrap="nowrap">Parcel Id:</td><td><%=sParcelid%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap">Use Type:</td><td><%=GetPermitUseType( iPermitId ) %></td></tr>
	<%			If PermitHasDetail( iPermitid, "useclassid" ) Then	%>
					<tr><td class="labelcell" nowrap="nowrap">Use Class:</td><td><%=GetPermitUseClass( iPermitId ) %></td></tr>
	<%			End If	
				If PermitHasDetail( iPermitid, "descriptionofwork" ) Then	%>
					<tr><td class="labelcell" nowrap="nowrap">Description of Work:</td><td><%=sDescriptionOfWork%></td></tr>
	<%			End If	
				If PermitHasDetail( iPermitid, "workclass" ) Then	%>
					<tr><td class="labelcell" nowrap="nowrap">Work Class:</td><td><%=GetPermitWorkClass( iPermitId ) %></td></tr>
	<%			End If	
				If PermitHasDetail( iPermitid, "workscope" ) Then	%>
					<tr><td class="labelcell" nowrap="nowrap">Work Scope:</td><td><%=GetPermitWorkScope( iPermitId ) %></td></tr>
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
	<%								End If	
				If PermitHasDetail( iPermitid, "zoning" ) Then	%>
					<tr><td class="labelcell" nowrap="nowrap">Zoning:</td><td><%=sZoning%></td></tr>
	<%								End If		
				If PermitHasDetail( iPermitid, "planno" ) Then	%>
					<tr><td class="labelcell" nowrap="nowrap">Plan #:</td><td><%=sPlanNumber%></td></tr>
	<%								End If		
				If PermitHasDetail( iPermitid, "demolishexisting" ) Then	
					If sDemolishExistingStructure <> "" Then %>
						<tr><td class="labelcell" nowrap="nowrap">&nbsp;</td><td>
							Existing structure will be demolished
						</td></tr>
						<% End If		%>

	<%								End If		
				If PermitHasDetail( iPermitid, "landfill" ) Then	%>
					<tr><td class="labelcell" nowrap="nowrap">Landfill Name:</td><td><%=sLandFillName%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap">Landfill City:</td><td><%=sLandFillCity%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap">Landfill Phone:</td><td><%=sLandFillPhone%></td></tr>
	<%								End If	
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
				Else %>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Temporary CO:</td><td>Does not have details</td></tr>
	<%			End If	
				If PermitHasDetail( iPermitid, "conotes" ) Then	
					If bHasCO Then		%>
						<tr><td class="labelcell" nowrap="nowrap" valign="top" colspan="2">Cert of Occupancy Stipulations, Conditions, Variances:</td></tr>
						<tr><td class="COcell" nowrap="nowrap" valign="top" colspan="2"><%=sCONotes%></td></tr>
	<%				End If			
				End If		%>
				</table>
		</p>
		<p>

				<span class="keyinfo">Contacts</span><br />
				<table cellpadding="2" border="0" cellspacing="0" class="viewdetails">
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Applicant:</td><td><%=GetPermitContactDetails( iPermitId, "isapplicant" )%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Primary Contact:</td><td><%=sPrimaryContact%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Billing Contact:</td><td><%=GetPermitContactDetails( iPermitId, "isbillingcontact" )%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Primary Contractor:</td><td><%=GetPermitContactDetails( iPermitId, "isprimarycontractor" )%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Architect/Engineer:</td><td><%=GetPermitContactDetails( iPermitId, "isarchitect" )%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Plans By:</td><td><%=GetPermitPlansByContact( iPermitId )%></td></tr>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Other Contractors:</td><td><%=GetPermitContactDetails( iPermitId, "iscontractor" )%></td></tr>
				</table>
		</p>
		<p><br style="page-break-before: always;" clear="all" /></p>
		<p>

				<span class="keyinfo">Fees</span><br />
				<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
					<tr><th class="firstcell">Finished Sq Ft</th><th>Unfinished Sq Ft</th><th>Total Sq Ft</th><th>Other Sq Ft</th><th>Job Value</th><th>Examination<br />Hours</th><th>Fee Total</th></tr>
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
				<table cellpadding="2" border="0" cellspacing="0" class="viewdetails">
					<tr><th class="firstcell">Category</th><th>Description</th><th>Method</th><th>Fee Amount</th></tr>
	<%				ShowPermitFees iPermitId 		%>	
				</table>

		</p>
		<p>
				<span class="keyinfo">Invoices &amp; Payments</span><br />
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

		</p>
		<p>

				<span class="keyinfo">Reviews</span><br />
				<p>
					<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
					<tr><th class="firstcell">Review</th><th>Status</th><th>Date</th><th>Reviewer</th></tr>
	<%				ShowReviewList iPermitId 		%>		
				</table>

		</p>
		<p><br style="page-break-before: always;" clear="all" /></p>
		<p>

				<span class="keyinfo">Inspections</span><br />
				<p>
					<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
					<tr><th class="firstcell">Inspection</th><th>Reinspection</th><th>Status</th><th>Inspected<br />Date</th><th>Inspector</th></tr>
	<%				ShowInspectionList iPermitId 		%>									
				</table>

		</p>
		<p>

				<span class="keyinfo">Attachments</span><br />
				<p>
					<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
	<%				ShowAttachmentList iPermitId 		%>		
				</table>

		</p>
		<p>

				<span class="keyinfo">Permit Notes</span><br />
				<p>
					<table cellpadding="4" cellspacing="0" border="0" class="viewdetails">
	<%				ShowPermitNotes iPermitId		%>
				</table>
		</p>
<%	Else		%>
		</div>
		
		<p>The requested permit could not be found, or you do not have permission to view it.</p>

<%	End If		%>

		<p id="pageendbuffer">&nbsp;</p>

	</div>
</div>

<!--END: PAGE CONTENT-->

<!--SPACING CODE-->
<p><br />&nbsp;<br />&nbsp;</p>
<!--SPACING CODE-->

<!--#Include file="../include_bottom.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' boolean GetPermitDetails( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetPermitDetails( ByVal iPermitId )
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
	sSql = sSql & " ISNULL(landfillname,'') AS landfillname, ISNULL(landfillcity,'') AS landfillcity, ISNULL(landfillphone,'') AS landfillphone "
	sSql = sSql & " FROM egov_permits WHERE permitid = " & iPermitId & " AND orgid = " & iOrgId 

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
			sDemolishExistingStructure = "X"
		Else
			sDemolishExistingStructure = ""
		End If 
		sLandFillName = oRs("landfillname")
		sLandFillCity = oRs("landfillcity")
		sLandFillPhone = oRs("landfillphone")
		GetPermitDetails = True 
	Else
		GetPermitDetails = False 
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 




%>
