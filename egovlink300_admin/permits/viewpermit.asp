<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewpermit.asp
' AUTHOR: Steve Loar
' CREATED: 05/19/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module displays one invoice for a permit.
'
' MODIFICATION HISTORY
' 1.0   05/19/2008	Steve Loar - Initial Version
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sPermitLocation, sLegalDescription, sListedOwner, iPermitAddressId, sListFixtures
Dim sShowConstructionType, sShowFeeTotal, sShowOccupancyType, sShowJobValue, sShowWorkDesc
Dim sShowFootages, sShowProposedUse, sPermitNotes, bListFixtures, bShowConstructionType, bShowFeeTotal
Dim bShowOccupancyType, bShowJobValue, bShowWorkDesc, bShowFootages, bShowProposedUse, bShowOtherContacts
Dim sShowElectricalContractor, sShowMechanicalContractor, sShowPlumbingContractor, sShowApplicantLicense
Dim iPrimaryContactId, bShowCounty, bShowParcelid, bShowPlansBy, sCounty, sParcelid, iPlansByContactId
Dim oAddressOrg, bShowPrimaryContact, bShowTotalSqFt, bShowApprovedAs, bShowFeeTypeTotals, bShowOccupancyUse
Dim sValue, sLabel, sApprovedAs, dZoneFee, dBBSFee, dRecImpactFee, dWaterImpactFee, dRoadImpactFee, dPermitFee
Dim sPrimaryContact, bShowPayments

sLevel = "../" ' Override of value from common.asp

Set oAddressOrg = New classOrganization 

iPermitId = CLng(request("permitid"))

sPermitLocation = GetPermitLocation( iPermitId, sLegalDescription, sListedOwner, iPermitAddressId, sCounty, sParcelid, True )

GetPermitDocumentShowFlags iPermitId, bListFixtures, bShowConstructionType, bShowFeeTotal, bShowOccupancyType, bShowJobValue, bShowWorkDesc, bShowFootages, bShowProposedUse, bShowOtherContacts, sShowElectricalContractor, sShowMechanicalContractor, sShowPlumbingContractor, sShowApplicantLicense, bShowCounty, bShowParcelid, bShowPlansBy, bShowPrimaryContact, bShowTotalSqFt, bShowApprovedAs, bShowFeeTypeTotals, bShowOccupancyUse, bShowPayments

sPermitNotes = GetPermitNotes( iPermitId )

sPrimaryContact = GetPermitDetailItemAsString( iPermitId, "primarycontact" )
'ShowPrimaryContactForPermit( iPermitId )
If sPrimaryContact = "" Then
	sPrimaryContact = "&nbsp;"
End If 

%>

<html>
<head>
	<title>E-Gov Permit</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitprint.css" media="print" />

	<script language="Javascript">
	<!--
		
		function doClose()
		{
			window.close();
			window.opener.focus();
		}

	//-->
	</script>

</head>

<body id="permitbody">
 
<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="Close" onclick="doClose();" /> 
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
<%	ShowPermitHeader iPermitId 	%>
	
	<table cellpadding="4" cellspacing="0" border="0" id="permitbody">
		<tr>
			<td class="permitbodylabel" nowrap="nowrap" valign="top">Job Site Address:</td><td valign="top" nowrap="nowrap"><%=sPermitLocation%></td>
			<td class="permitbodylabel" nowrap="nowrap" valign="top">Legal Description:</td><td valign="top"><%=sLegalDescription%></td>
		</tr>
<%		If bShowCounty Or bShowParcelid Then %>
		<tr>
<%			If bShowCounty Then		%>
				<td class="permitbodylabel" nowrap="nowrap"><%=oAddressOrg.GetOrgDisplayName("address grouping field")%>:</td><td nowrap="nowrap"><%=sCounty%></td>
<%			Else 	%>	
				<td colspan="2">&nbsp;</td>
<%			End If	%>	
<%			If bShowParcelid Then		%>
				<td class="permitbodylabel" nowrap="nowrap">Parcel Id:</td><td><%=sParcelid%></td>
<%			Else 	%>	
				<td colspan="2">&nbsp;</td>
<%			End If	%>	
		</tr>
<%		End If	%>	
		<tr>
			<td class="permitbodylabel" nowrap="nowrap" valign="top">Primary Contractor:</td><td valign="top" nowrap="nowrap"><%=ShowPrimaryContractorForPermit( iPermitId )%></td>
			<td class="permitbodylabel" nowrap="nowrap" valign="top">Property Owner:</td><td colspan="3" valign="top"><%=sListedOwner%></td>
		</tr>
<%		If bShowPrimaryContact Then  %>
		<tr>
			<td class="permitbodylabel" nowrap="nowrap" valign="top">Primary Contact:</td><td valign="top" nowrap="nowrap"><%=sPrimaryContact %></td>
		</tr>
<%		End If	%>

<%		If sShowApplicantLicense Then	
			iPrimaryContactId = GetPrimaryContactIdForPermit( iPermitId )
			ShowPrimaryContactLicense iPermitId, iPrimaryContactId 
		End If		%>

<%		'If bShowOtherContacts Then	
		'	ShowOtherContacts iPermitId
		'End If		%>

<%		If sShowPlumbingContractor Then	
			ShowContractor iPermitId, "isplumbing"
		End If		%>
<%		If sShowMechanicalContractor Then	
			ShowContractor iPermitId, "ismechanical"
		End If		%>
<%		If sShowElectricalContractor Then	
			ShowContractor iPermitId, "iselectrical"
		End If		%>

<%		If bShowPlansBy Then  
			ShowPlansByForPermit iPermitId 
		End If		%>

<%		If bShowApprovedAs Then	
			sApprovedAs = GetPermitDetailItemAsString( iPermitId, "approvedas" )
			If sApprovedAs = "" Then
				sApprovedAs = "&nbsp;"
			End If 
%>
			<tr class="permitsectionstart">
				<td class="permitbodylabel" nowrap="nowrap">Permit is Granted For:</td><td colspan="3"><%= sApprovedAs %></td>
			</tr>
<%		End If		%>

<%		If bShowWorkDesc Then 
			sDescriptionOfWork = GetPermitDetailItemAsString( iPermitId, "descriptionofwork" )
			If sDescriptionOfWork = "" Then
				sDescriptionOfWork = "&nbsp;"
			End If 
%>
		<tr class="permitsectionstart">
			<td class="permitbodylabel" nowrap="nowrap">Description of work:</td><td colspan="3"><%= sDescriptionOfWork %></td>
		</tr>
<%		End If	%>


<%		If bShowConstructionType Or bShowOccupancyType Or bShowOccupancyUse Then %>
			<tr class="permitsectionstart">

<%			If bShowConstructionType Then 
				sConstructionType = GetPermitConstructionType( iPermitId )	
				If sConstructionType = "" Then 
					sConstructionType = "&nbsp;"
				End If 
%>
				<td class="permitbodylabel" nowrap="nowrap" valign="top">Construction Type:</td><td valign="top"><%= sConstructionType %></td>
<%			Else 	%>	
				<td colspan="2">&nbsp;</td>
<%			End If	%>	

<%			If bShowOccupancyType Or bShowOccupancyUse Then %>
				
<%				If bShowOccupancyUse Then 
					sValue = GetPermitOccupancyTypeGroup( iPermitId )
					If sValue = "" Then 
						sValue = "&nbsp;"
					End If
					sLabel = "Use Group:"
				End If %>
<%				If bShowOccupancyType Then 
					sOccupancy = GetPermitOccupancyType( iPermitId )
					If sOccupancy = "" Then 
						sOccupancy = "&nbsp;"
					End If
					If sLabel <> "" Then 
						sLabel = sLabel & "<br />"
					End If 
					sLabel = sLabel & "Occupancy Type:"
					If sValue <> "" Then 
						sValue = sValue & "<br />"
					End If 
					sValue = sValue & sOccupancy
				End If 
%>
				<td class="permitbodylabel" nowrap="nowrap" valign="top"><%=sLabel%></td><td valign="top"><%=sValue%></td>
<%			Else 	%>	
				<td colspan="2">&nbsp;</td>
<%			End If	%>
			</tr>
<%		End If	%>

<%		If bShowProposedUse Then	
			sProposedUse = GetPermitDetailItemAsString( iPermitId, "proposeduse" )
			If sProposedUse = "" Then
				sProposedUse = "&nbsp;"
			End If 
%>
			<tr class="permitsectionstart">
				<td class="permitbodylabel" nowrap="nowrap">Proposed use:</td><td colspan="3"><%= sProposedUse %></td>
			</tr>
<%		End If		%>

<%		If bShowFootages Then	%>
			<tr class="permitsectionstart">
				<td class="permitbodylabel" nowrap="nowrap">Finished Sq Feet:</td><td><%=GetPermitDetailItemAsNumber( iPermitId, "finishedsqft", "integer" )%></td>
				<td class="permitbodylabel" nowrap="nowrap">Unfinished Sq Ft:</td><td><%=GetPermitDetailItemAsNumber( iPermitId, "unfinishedsqft", "integer" )%></td>
			</tr>
			<tr>
				<td class="permitbodylabel" nowrap="nowrap">Total Sq Feet:</td><td><%=GetPermitDetailItemAsNumber( iPermitId, "totalsqft", "integer" )%></td>
				<td class="permitbodylabel" nowrap="nowrap">Other Sq Ft:</td><td><%=GetPermitDetailItemAsNumber( iPermitId, "othersqft", "integer" )%></td>
			</tr>
<%		End If		%>

<%		If bShowTotalSqFt Then	%>
			<tr class="permitsectionstart">
				<td class="permitbodylabel" nowrap="nowrap" colspan="2">Total&nbsp;Floor&nbsp;Area&nbsp;Exterior&nbsp;Dimensions:</td><td colspan="2"><%=FormatNumber(GetPermitDetailItemAsNumber( iPermitId, "totalsqft", "integer" ),0)%> sq.ft.</td>
			</tr>
<%		End If					%>

<%		If bShowJobValue Or bShowFeeTotal Then %>
			<tr class="permitsectionstart">
<%			If bShowJobValue Then %>
				<td class="permitbodylabel" nowrap="nowrap">Total Valuation:</td><td><%=GetPermitDetailItemAsNumber( iPermitId, "jobvalue", "currency" )%></td>
<%			Else 	%>	
				<td colspan="2">&nbsp;</td>
<%			End If	%>	
<%			If bShowFeeTotal Then %>
				<td class="permitbodylabel" nowrap="nowrap">Total Fees:</td><td><%=GetPermitDetailItemAsNumber( iPermitId, "feetotal", "currency" )%></td>
<%			Else 	%>	
				<td colspan="2">&nbsp;</td>
<%			End If	%>
			</tr>
<%		End If	%>
	</table>

<%	If bShowFeeTypeTotals Then	
		dPermitFee = GetPermitBuildingFees( iPermitId )
		If CDbl(dPermitFee) = CDbl(0.00) Then
			dPermitFee = " &nbsp; &nbsp; &nbsp; &nbsp; "
		Else
			dPermitFee = " &nbsp;" & FormatCurrency(dPermitFee,2) & " &nbsp; "
		End If 
		dZoneFee = GetPermitFeeTypeTotal( iPermitId, "iszone" )
		If CDbl(dZoneFee) = CDbl(0.00) Then
			dZoneFee = " &nbsp; &nbsp; &nbsp; &nbsp; "
		Else
			dZoneFee = " &nbsp;" & FormatCurrency(dZoneFee,2) & " &nbsp; "
		End If 
		dBBSFee = GetPermitFeeTypeTotal( iPermitId, "isbbs" )
		If CDbl(dBBSFee) = CDbl(0.00) Then
			dBBSFee = " &nbsp; &nbsp; &nbsp; &nbsp; "
		Else
			dBBSFee = " &nbsp;" & FormatCurrency(dBBSFee,2) & " &nbsp; "
		End If 
		dRecImpactFee = GetPermitFeeTypeTotal( iPermitId, "isrecreationimpact" )
		If CDbl(dRecImpactFee) = CDbl(0.00) Then
			dRecImpactFee = " &nbsp; &nbsp; &nbsp; &nbsp; "
		Else
			dRecImpactFee = " &nbsp;" & FormatCurrency(dRecImpactFee,2) &" &nbsp; "
		End If 
		dWaterImpactFee = GetPermitFeeTypeTotal( iPermitId, "iswatermeter" )
		If CDbl(dWaterImpactFee) = CDbl(0.00) Then
			dWaterImpactFee = " &nbsp; &nbsp; &nbsp; &nbsp; "
		Else
			dWaterImpactFee = " &nbsp;" & FormatCurrency(dWaterImpactFee,2) & " &nbsp; "
		End If 
		dRoadImpactFee = GetPermitFeeTypeTotal( iPermitId, "isroadimpact" )
		If CDbl(dRoadImpactFee) = CDbl(0.00) Then
			dRoadImpactFee = " &nbsp; &nbsp; &nbsp; &nbsp; "
		Else
			dRoadImpactFee = " &nbsp;" & FormatCurrency(dRoadImpactFee,2) & " &nbsp; "
		End If 
%>
	<p id="feetypetotals" class="viewpermit">
		<strong>FEES:</strong><br />
		Permit:<span class="feetypetotal"><%=dPermitFee%></span> State Fees:<span class="feetypetotal"><%=dBBSFee%></span> Zoning Fees:<span class="feetypetotal"><%=dZoneFee%></span><br />
		<strong>Other Fees:</strong><br />
		Road Impact:<span class="feetypetotal"><%=dRoadImpactFee%></span> Water Impact:<span class="feetypetotal"><%=dWaterImpactFee%></span> Rec. Impact:<span class="feetypetotal"><%=dRecImpactFee%></span> 
	</p>
<%	End If							%>

<%	If bShowPayments Then			%>
		<p class="viewpermit">
			<table cellpadding="2" cellspacing="0" border="0" class="viewdetails" id="permitpayments">
				<tr><th class="firstcell">Payment #</th><th>Date</th><th>Method</th><th>Amount<br />Paid/Waived</th></tr>
				<% ShowPayments iPermitId %>
			</table>
		</p>
<%	End If							%>

<%	If bListFixtures Then 
		ShowPermitFixtures iPermitId
	End If	%>

<%	If sPermitNotes <> "" Then	%>
	<p class="viewpermit">
		<span id="permitnotes">Additional Comments:&nbsp;<%=sPermitNotes%></span>
	</p>
<%	End If		%>

<%	ShowPermitFooter iPermitId 	%>

	</div>
</div>
<!--END: PAGE CONTENT-->


</body>
</html>


<%
Set oAddressOrg = Nothing 
'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void ShowPermitHeader( )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitHeader( ByVal iPermitId )
	Dim sPermitLogo

	response.write vbcrlf & "<table id=""permitheader"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
	response.write vbcrlf & "<tr>"

	sPermitLogo = GetPermitDocumentValue( iPermitId, "permitlogo" )
	If sPermitLogo <> "" Then
		response.write "<td><img src=""" & sPermitLogo & """ alt=""logo"" border=""0"" /></td>"
	Else
		response.write "<td>&nbsp;</td>"
	End If 

	' Center title including title and sub title
	response.write "<td align=""center"" valign=""top"">"
	response.write "<span id=""permittitle"">" & GetPermitDocumentValue( iPermitId, "permittitle" ) & "</span><br />"
	response.write "<span id=""permitsubtitle"">" & GetPermitDocumentValue( iPermitId, "permitsubtitle" ) & "</span>"
	response.write "</td>"

	' Permit Number and any right titles
	response.write "<td valign=""top"" nowrap=""nowrap""><span id=""permitnolabel"">Permit No: </span>"
	response.write "<span id=""permitno"">" & GetPermitNumber( iPermitId ) & "</span><br /><br />"
	response.write "<span id=""permitrighttitle"">" & GetPermitDocumentValue( iPermitId, "permitrighttitle" ) & "</span>"
	response.write "</td>"

	response.write "</tr>"

	' spacer row
	response.write vbcrlf & "<tr><td colspan=""3"">&nbsp;</td></tr>"

	' Title Bottom
	response.write vbcrlf & "<tr>"
	response.write "<td colspan=""3"" align=""right""><span id=""permittitlebottom"">"
	response.write GetPermitDocumentValue( iPermitId, "permittitlebottom" )
	response.write "</span></td>"
	response.write "</tr>"
	response.write vbcrlf & "</table>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPermitFooter( iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFooter( iPermitId )
	Dim sApprovedBy

	response.write vbcrlf & "<p class=""viewpermit"">"
	response.write vbcrlf & GetPermitDocumentValue( iPermitId, "additionalfooterinfo" )
	response.write vbcrlf & "</p>"

	response.write vbcrlf & "<p id=""permitfooter"" class=""viewpermit"">"
	response.write vbcrlf & GetPermitDocumentValue( iPermitId, "permitfooter" )
	response.write vbcrlf & "</p>"

	response.write vbcrlf & "<p id=""permitsubfooter"" class=""viewpermit"">"
	response.write vbcrlf & GetPermitDocumentValue( iPermitId, "permitsubfooter" )
	response.write vbcrlf & "</p>"

	response.write vbcrlf & vbcrlf & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" id=""permitsignatureline"">"
	response.write vbcrlf & "<tr>"
	response.write "<td valign=""top"" nowrap=""nowrap"" class=""approvingofficial""><span class=""approvingofficial"">Approving Official: </span></td>"
	response.write "<td valign=""top"" nowrap=""nowrap""><span class=""approvingofficial"">" & GetPermitDocumentValue( iPermitId, "approvingofficial" ) & "</span></td>"
	response.write "<td valign=""top"" align=""right"" nowrap=""nowrap""><span id=""issueddate"">Issued Date: " & GetPermitIssuedDate( iPermitId ) & "</span></td>"
	response.write "</tr>"
	response.write vbcrlf & "</table>"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPermitFixtures( iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowPermitFixtures( iPermitId )
	Dim sSql, oRs, iPermitFeeId

	iPermitFeeId = CLng(0)

	sSql = "SELECT P.permitfeeid, ISNULL(P.permitfeeprefix,'') AS permitfeeprefix, ISNULL(P.permitfee,'') AS permitfee, "
	sSql = sSql & " ISNULL(F.permitfixture,'') AS permitfixture, ISNULL(F.qty,0) AS qty "
	sSql = sSql & " FROM egov_permitfees P, egov_permitfixtures F "
	sSql = sSql & " WHERE F.permitid = P.permitid AND F.permitfeeid = P.permitfeeid AND "
	sSql = sSql & " F.isincluded = 1 AND P.includefee = 1 and P.permitid = " & iPermitId
	sSql = sSql & " ORDER BY F.displayorder, F.permitfixture"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<table cellpadding=""3"" cellspacing=""0"" border=""0"" id=""permitfixturelist"">"
		Do While Not oRs.EOF
			If iPermitFeeId <> CLng(oRs("permitfeeid")) Then
				iPermitFeeId = CLng(oRs("permitfeeid"))
				response.write vbcrlf & "<tr><td colspan=""2""><strong>"
				If oRs("permitfeeprefix") <> "" Then 
					response.write oRs("permitfeeprefix") & " - "
				End If 
				response.write oRs("permitfee") & "</strong></td></tr>"
				response.write vbcrlf & "<tr><td><strong>Fixture</strong></td><td><strong>Quantity</strong></td></tr>"
			End If 
			response.write vbcrlf & "<tr><td>" & oRs("permitfixture") & "</td><td>" & oRs("qty") & "</td></tr>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowOtherContacts( iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowOtherContacts( iPermitId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0 
	sSql = "SELECT ISNULL(contractortypeid,0) AS contractortypeid, company, lastname, firstname, phone "
	sSql = sSql & " FROM egov_permitcontacts WHERE iscontractor = 1 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		Do While Not oRs.EOF 
			iRowCount = iRowCount + 1
			If CLng(iRowCount) = CLng(1) Then 
				response.write vbcrlf & "<tr class=""permitsectionstart"">"
			Else
				response.write vbcrlf & "<tr>"
			End If 
			sContractorType = GetContractorType( oRs("contractortypeid") )
			response.write "<td class=""permitbodylabel"" nowrap=""nowrap"" valign=""top"">" & sContractorType
			If sContractorType <> "" Then
				response.write ":"
			End If 
			response.write "</td>"
			response.write "<td valign=""top"">"
			If oRs("firstname") <> "" Then
				response.write oRs("firstname") & " " & oRs("lastname")
			End If 
			If oRs("company") <> "" Then
				If oRs("firstname") <> "" Then 
					response.write "<br />( " & oRs("company") & " ) "
				Else
					response.write oRs("company")
				End If 
			End If 
			response.write "</td>"
			response.write "<td colspan=""3"" valign=""top"">"
			If Not IsNull(oRs("phone")) And oRs("phone") <> "" Then
				response.write FormatPhoneNumber( oRs("phone") )
			Else
				response.write "&nbsp;"
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
' Sub ShowContractor( iPermitId, sContractorTypeFlag )
'--------------------------------------------------------------------------------------------------
Sub ShowContractor( iPermitId, sContractorTypeFlag )
	Dim sSql, oRs

	iRowCount = 0 
	sSql = "SELECT P.company, P.lastname, P.firstname, P.phone, C.contractortype "
	sSql = sSql & " FROM egov_permitcontacts P, egov_permitcontractortypes C "
	sSql = sSql & " WHERE P.contractortypeid = C.contractortypeid AND " & sContractorTypeFlag & " = 1 "
	sSql = sSql & " AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<tr class=""permitsectionstart"">"
		sContractorType = oRs("contractortype") & " Contractor"
		response.write "<td class=""permitbodylabel"" nowrap=""nowrap"" valign=""top"">" & sContractorType
		If sContractorType <> "" Then
			response.write ":"
		End If 
		response.write "</td>"
		response.write "<td valign=""top"">"
		If oRs("firstname") <> "" Then
			response.write oRs("firstname") & " " & oRs("lastname")
		End If 
		If oRs("company") <> "" Then
			If oRs("firstname") <> "" Then 
				response.write "<br />( " & oRs("company") & " ) "
			Else
				response.write oRs("company")
			End If 
		End If 
		response.write "</td>"
		response.write "<td colspan=""3"" valign=""top"">"
		If Not IsNull(oRs("phone")) And oRs("phone") <> "" Then
			response.write FormatPhoneNumber( oRs("phone") )
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPrimaryContactLicense( iPermitId, iPermitContactId )
'--------------------------------------------------------------------------------------------------
Sub ShowPrimaryContactLicense( iPermitId, iPermitContactId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(licensetype,'') AS licensetype, licenseenddate, ISNULL(licensenumber,'&nbsp;') AS licensenumber "
	sSql = sSql & " FROM egov_permitcontacts_licenses WHERE permitid = " & iPermitID
	sSql = sSql & " AND permitcontactid = " & iPermitContactId & " ORDER BY licenseenddate DESC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<tr class=""permitsectionstart"">"
		response.write "<td class=""permitbodylabel"">"
		response.write "<strong>" & oRs("licensetype") & " License: </strong>"
		response.write "</td>"
		response.write "<td>" & oRs("licensenumber") & "</td>"
		response.write "<td>"
		response.write "<strong>Expiration Date: </strong>"
		response.write "</td>"
		response.write "<td>"
		If Not IsNull(oRs("licenseenddate")) Then
			response.write FormatDateTime(oRs("licenseenddate"),2)
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td>"
		response.write "</tr>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPlansByForPermit( iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowPlansByForPermit( iPermitId )
	Dim sSql, oRs, sResults

	sResults = ""

	sSql = "SELECT C.company, C.firstname, C.lastname, C.phone "
	sSql = sSql & " FROM egov_permitcontacts C, egov_permits P "
	sSql = sSql & " WHERE C.permitcontactid = P.plansbycontactid AND P.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<tr class=""permitsectionstart"">"
	response.write "<td class=""permitbodylabel"" nowrap=""nowrap"">Plans By:</td>"
	response.write "<td>"

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then
			response.write oRs("firstname") & " " & oRs("lastname")
		End If 
		If oRs("company") <> "" Then
			If oRs("firstname") <> "" Then 
				response.write "<br />( " & oRs("company") & " ) "
			Else
				response.write oRs("company")
			End If 
		End If 
		response.write "</td>"
		response.write "<td colspan=""3"" valign=""top"">"
		If Not IsNull(oRs("phone")) And oRs("phone") <> "" Then
			response.write FormatPhoneNumber( oRs("phone") )
		Else
			response.write "&nbsp;"
		End If 
	Else
		response.write "&nbsp;"
		response.write "</td>"
		response.write "<td colspan=""3"">&nbsp;"
	End If 

	response.write "</td>"
	response.write "</tr>"

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPayments( iPermitId )
'--------------------------------------------------------------------------------------------------
Sub ShowPayments( iPermitId )
	Dim sSql, oRs, dTotal

	dTotal = CDbl(0.00) 

	sSql = "SELECT ISNULL(SUM(L.amount),0.00) AS paymenttotal, L.paymentid, J.paymentdate "
	sSql = sSql & " FROM egov_accounts_ledger L, egov_class_payment J "
	sSql = sSql & " WHERE L.paymentid = J.paymentid AND L.ispaymentaccount = 1 AND L.permitid = " & iPermitId
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
' Sub ShowInvoicePayments( iPaymentId )
'--------------------------------------------------------------------------------------------------
Sub ShowInvoicePayments( iPaymentId )
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





%>
