<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: viewpermitarchive.asp
' AUTHOR: Steve Loar
' CREATED: 02/08/2011
' COPYRIGHT: Copyright 2011 eclink, inc.
'			 All Rights Reserved.
'
' Description:  displays archived permit details. Taken from viewpermitdetails.asp
'
' MODIFICATION HISTORY
' 1.0   02/08/2011	Steve Loar - INITIAL VERSION
' 1.1	05/13/2011	Steve Loar - Changes added for Camden County
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql, oRs, sApplied, sReleased, sApproved, sIssued, sCompleted, sExpires

iPermitId = CLng(request("permitid"))

%>

<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="permits.css" />
	<link rel="stylesheet" type="text/css" href="permitdetailsprint.css" media="print" />

	<script type="text/javascript" src="../scripts/layers.js"></script>

	<script language="Javascript">
	<!--

		function doBack()
		{
			history.go(-1);
		}

	//-->
	</script>

</head>

<body>

<div id="idControls" class="noprint">
	<input type="button" class="button ui-button ui-widget ui-corner-all" value="<< Back" onclick="doBack();" /> &nbsp;&nbsp;
	<input type="button" class="button ui-button ui-widget ui-corner-all" onclick="javascript:window.print();" value="Print" />&nbsp;&nbsp;
</div>

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	
		<!--BEGIN: PAGE TITLE-->
		<p>
			<font size="+1"><strong>Permit Details</strong></font><br /><br />
		</p>
		<!--END: PAGE TITLE-->
<%

	sSql = "SELECT actualpermitnumber, permittypedesc, ISNULL(descriptionofwork,'') AS descriptionofwork, ISNULL(proposeduse,'') AS proposeduse, "
	sSql = sSql & "applieddate, releaseddate, approveddate, issueddate, completeddate, expirationdate, ISNULL(permitstatus,'') AS permitstatus, "
	sSql = sSql & "ISNULL(workclass,'') AS workclass, ISNULL(totalpaid,0.00) AS totalpaid, ISNULL(jobvalue,0.00) AS jobvalue, "
	sSql = sSql & "ISNULL(listedowner,'') AS listedowner, ISNULL(owneraddress,'') AS owneraddress, ISNULL(ownercity,'') AS ownercity, "
	sSql = sSql & "ISNULL(ownerstate,'') AS ownerstate, ISNULL(ownerzip,'') AS ownerzip, ISNULL(ownerphone,'') AS ownerphone, "
	sSql = sSql & "ISNULL(contractorcompany,'') AS contractorcompany, ISNULL(contractorname,'') AS contractorname, "
	sSql = sSql & "ISNULL(contractoraddress,'') AS contractoraddress, ISNULL(contractorpobox,'') AS contractorpobox, "
	sSql = sSql & "ISNULL(contractorcity,'') AS contractorcity, ISNULL(contractorstate,'') AS contractorstate, "
	sSql = sSql & "ISNULL(contractorzip,'') AS contractorzip, ISNULL(contractorphone,'') AS contractorphone, ISNULL(contractorlicense,'') AS contractorlicense, "
	sSql = sSql & "ISNULL(jobaddress,'') AS jobaddress, ISNULL(residentunit,'') AS residentunit, ISNULL(residentcity,'') AS residentcity, "
	sSql = sSql & "ISNULL(residentstreetnumber,'') AS residentstreetnumber, ISNULL(residentstreetprefix,'') AS residentstreetprefix, "
	sSql = sSql & "ISNULL(residentstreetname, '') AS residentstreetname, ISNULL(streetsuffix,'') AS streetsuffix, ISNULL(streetdirection,'') AS streetdirection, "
	sSql = sSql & "ISNULL(residentstate,'') AS residentstate, ISNULL(residentzip,'') AS residentzip, ISNULL(parcelidnumber,'') AS parcelidnumber, "
	sSql = sSql & "ISNULL(propertynotes,'') AS propertynotes, ISNULL(floodzone,'') AS floodzone, ISNULL(township,'') AS township, "
	sSql = sSql & "ISNULL(zone,'') AS zone, ISNULL(udonumber,'') AS udonumber, ISNULL(bathrooms,0) AS bathrooms, ISNULL(bedrooms,0) AS bedrooms, "
	sSql = sSql & "ISNULL(fireplaces,0) AS fireplaces, ISNULL(units,0) AS units, ISNULL(totalaccsqft,0) AS totalaccsqft, "
	sSql = sSql & "ISNULL(totalfinishedsqft,0) AS totalfinishedsqft, ISNULL(totalfinunfinaccsqft,0) AS totalfinunfinaccsqft, "
	sSql = sSql & "ISNULL(totalunfinishedsqft,0) AS totalunfinishedsqft, ISNULL(occupancyinfo,'') AS occupancyinfo, "
	sSql = sSql & "ISNULL(parcellot,'') AS parcellot, ISNULL(subdivision,'') AS subdivision "
	sSql = sSql & "FROM egov_permitarchives WHERE permitid = " & iPermitId & " AND orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
%>

		<p>
			Permit #: <span class="keyinfo"><%=oRs("actualpermitnumber")%> &nbsp; &nbsp; &mdash; &nbsp; <%=oRs("permittypedesc") %></span>
		</p>

		<p>
			Permit Status: <span class="keyinfo"><%=oRs("permitstatus")%></span>
		</p>


<%
		If Not IsNull(oRs("applieddate")) Then
			sApplied = DateValue(oRs("applieddate"))
		Else
			sApplied = ""
		End If 
		If Not IsNull(oRs("releaseddate")) Then
			sReleased = DateValue(oRs("releaseddate"))
		Else
			sReleased = ""
		End If 
		If Not IsNull(oRs("approveddate")) Then
			sApproved = DateValue(oRs("approveddate"))
		Else
			sApproved = ""
		End If 
		If Not IsNull(oRs("issueddate")) Then
			sIssued = DateValue(oRs("issueddate"))
		Else
			sIssued = ""
		End If 
		If Not IsNull(oRs("completeddate")) Then
			sCompleted = DateValue(oRs("completeddate"))
		Else
			sCompleted = ""
		End If 
		If Not IsNull(oRs("expirationdate")) Then
			' And sCompleted = ""
			sExpires = DateValue(oRs("expirationdate"))
		Else
			sExpires = ""
		End If 
%>
		<p>
			<fieldset>
				<legend><span class="keyinfo">Critical Dates</span></legend><br />
				<table cellpadding="2" cellspacing="0" border="0" class="viewdetails">
					<tr><th class="firstcell" align="center">Applied</th><th>Released</th><th>Approved</th><th>Permit<br />Issued</th><th>Completed</th><th>Expires</th></tr>
					<tr>
						<td align="center" class="firstcell"><span class="detaildata"><%=sApplied%></span></td>
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
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Job Site Address:</td>
						<td>
<%							'  This address display was specified by Camden Co.
							response.write "<strong>" & oRs("residentstreetnumber") & " " & oRs("residentunit") & " " & oRs("residentstreetname") & " " & oRs("streetsuffix") & " " & oRs("streetdirection") & "</strong><br />"
							response.write oRs("residentcity") & ", " & oRs("residentstate") & " " & oRs("residentzip")
%>
						</td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Parcel Lot:</td>
						<td><% =oRs("parcellot") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Subdivision:</td>
						<td><% =oRs("subdivision") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">UDO Number:</td>
						<td><% =oRs("udonumber") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Description of Work:</td>
						<td><% =oRs("descriptionofwork") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Work Class:</td>
						<td><% =oRs("workclass") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Proposed Use:</td>
						<td><% =oRs("proposeduse") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Job Value:</td>
						<td><% =FormatCurrency(oRs("jobvalue"),2) %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Fees Paid:</td>
						<td><% =FormatCurrency(oRs("totalpaid"),2) %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Parcel ID:</td>
						<td><% =oRs("parcelidnumber") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Property Notes:</td>
						<td><% =oRs("propertynotes") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Flood Zone:</td>
						<td><% =oRs("floodzone") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Township:</td>
						<td><% =oRs("township") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Property Zone:</td>
						<td><% =oRs("zone") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Occupancy Info:</td>
						<td><% =oRs("occupancyinfo") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Bathrooms:</td>
						<td><% =oRs("bathrooms") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Bedrooms:</td>
						<td><% =oRs("bedrooms") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Fireplaces:</td>
						<td><% =oRs("fireplaces") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Units:</td>
						<td><% =oRs("units") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Total Acc SqFt:</td>
						<td><% =oRs("totalaccsqft") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Total Finished SqFt:</td>
						<td><% =oRs("totalfinishedsqft") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Total Unfinished SqFt:</td>
						<td><% =oRs("totalunfinishedsqft") %></td>
					</tr>
					<tr><td nowrap="nowrap" class="labelcell" valign="top">Total Fin Unfin Acc SqFt:</td>
						<td><% =oRs("totalfinunfinaccsqft") %></td>
					</tr>
				</table>
			</fieldset>
		</p>

		<p>
			<fieldset>
				<legend><span class="keyinfo">Contacts</span></legend><br />
				<table cellpadding="2" border="0" cellspacing="0" class="viewdetails">
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Contractor:</td>
						<td>
<%
							If oRs("contractorname") <> "" Then 
								response.write "<strong>" & oRs("contractorname") & "</strong><br />"
							End If 
							If oRs("contractorcompany") <> "" Then
								response.write "<strong>" & oRs("contractorcompany") & "</strong><br />"
							End If 
							response.write oRs("contractoraddress") & "<br />"
							If oRs("contractorpobox") <> "" Then
								response.write "P.O. Box " & oRs("contractorpobox") & "<br />"
							End If 
							response.write oRs("contractorcity") & ", " & oRs("contractorstate") & " " & oRs("contractorzip")
							If oRs("contractorphone") <> "" Then
								response.write "<br />" & oRs("contractorphone")
							End If 
%>
						</td>
						<td valign="top">
<%
							If oRs("contractorlicense") <> "" Then 
								response.write "Contractor License: <strong>" & oRs("contractorlicense") & "</strong>"
							End If 
%>
						</td>
					</tr>
					<tr><td class="labelcell" nowrap="nowrap" valign="top">Owner:</td>
						<td colspan="2">
<%
							If oRs("listedowner") <> "" Then 
								response.write "<strong>" & oRs("listedowner") & "</strong><br />"
								response.write oRs("owneraddress") & "<br />"
								response.write oRs("ownercity") & ", " & oRs("ownerstate") & " " & oRs("ownerzip")
								If oRs("ownerphone") <> "" Then
									response.write "<br />" & oRs("ownerphone")
								End If 
							End If
%>
						</td>
					</tr>
				</table>
			</fieldset>
		</p>

<%
	Else
%>
		<p>
			<span class="keyinfo">No information could be found for the requested permit.</span>
		</p>
<%

	End If 

	oRs.Close
	Set oRs = Nothing 

%>


	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>

</html>
