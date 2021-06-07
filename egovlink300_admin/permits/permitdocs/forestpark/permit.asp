<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
	<style>
	.row
	{
		width:100%;
	}
	.row div
	{
		display:inline-block;
		text-align:center;
	}
	.flexrow
	{
		width:100%;
		display:flex;
		flex-direction:row;
		flex-wrap:nowrap;
		justify-content:space-between;
	}
	.flexrow .value
	{
		flex-grow:4;
		border-bottom:1px solid black;
	}
	.flexrow .value.no-u
	{
		border-bottom:0;
	}
	h3
	{
		margin:0;
	}
	.bdrbtm td
	{
		border-bottom:1px solid black;
	}
	.bdrbtm u, .bold
	{
		font-weight:bold;
	}
	</style>
	</head>
	<%
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<!--body onLoad="window.print();"-->
	<body>
		<div class="row">
			<div style="width:19%;">LOGO</div>
			<div style="width:60%;">
				<h3>Forest Park</h3>
				Building & Safety Division<br />
				Permit Receipt
			</div>
			<div style="width:19%;">
				<u>Project Number</u><br />
				<%=strPermitNumber%>
			</div>

		</div>
		<table style="width:100%" border="1" cellspacing="0" cellpadding="0">
			<tr>
				<td width="50%" valign="top">
					<h3><u>Job Address</u></h3>
					<div class="flexrow">
						<div>Street:</div>
						<div class="value">
						<%
							response.write oRs("residentstreetnumber")
							If oRs("residentstreetprefix") <> "" Then
								response.write " " & oRs("residentstreetprefix")
							End If
							response.write " " & oRs("residentstreetname")
							If oRs("streetsuffix") <> "" Then
								response.write " " & oRs("streetsuffix")
							End If
							If oRs("streetdirection") <> "" Then
								response.write " " & oRs("streetdirection")
							End If
							If oRs("residentunit") <> "" Then
								response.write ", " & oRs("residentunit")
							End If
							%>
						</div>
					</div>
					<div class="flexrow">
						<div>Address Desc:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Tract:</div>
						<div class="value no-u">&nbsp;</div>
						<div>Lot:</div>
						<div class="value no-u">&nbsp;</div>
						<div>APN:</div>
						<div class="value no-u">&nbsp;</div>
					</div>
					<hr noshade>
					<div class="flexrow">
						<div class="bold">Owner:</div>
						<div class="value"><%=strOwnerName%>&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Address:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>City/St/Zip:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Phone:</div>
						<div class="value no-u">&nbsp;</div>
						<div>Owner/Building Permit:</div>
						<div class="value no-u">&nbsp;</div>
					</div>
					<hr noshade>
					<div class="flexrow">
						<div class="bold">Applicant:</div>
						<div class="value"><%=oRs("appfirstname")%>&nbsp;<%=oRs("applastname")%>&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Address:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>City/St/Zip:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Phone:</div>
						<div class="value no-u">&nbsp;</div>
					</div>
					<hr noshade>
					<div class="flexrow">
						<div class="bold">Contractor:</div>
						<div class="value"><%=oRs("appfirstname")%>&nbsp;<%=oRs("applastname")%>&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Contact:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Address:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>City/St/Zip:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Phone:</div>
						<div class="value">&nbsp;</div>
						<div>Cell:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Lic. #:</div>
						<div class="value">&nbsp;</div>
						<div>Exp. Date:</div>
						<div class="value">&nbsp;</div>
						<div>Class:</div>
						<div class="value">&nbsp;</div>
					</div>
					<br />
					<center class="bold">Workers' Compensation Carrier</center>
					<br />
					<div class="flexrow">
						<div style="flex-grow:1;">
							<div class="value">&nbsp;</div>
							<div style="text-align:center;">Policy #</div>
						</div>
						<div style="flex-grow:1;">
							<div class="value">&nbsp;</div>
							<div style="text-align:center;">Exp. Date</div>
						</div>
					</div>
					<hr noshade>
					<center class="bold">Contractor to provide access to Roof &amp; Attic</center>
					<hr noshade>
					<center>
						SEAL
					</center>
				</td>
				<td valign="top" width="50%">
					<table width="100%" class="bdrbtm">
						<tr>
							<td align="center" style="border-right:1px solid black;">
								<u>Application Date</u>
								<br />
								<%=FormatDateTime(oRs("applieddate"),2)%>
							</td>
							<td align="center" colspan="2" style="border-right:1px solid black;">
								<u>Issued Date</u>
								<br />
								<%=FormatDateTime(oRs("issueddate"),2)%>
							</td>
							<td align="center">
								<u>Expired Date</u>
								<br />
								<%=FormatDateTime(oRs("expirationdate"),2)%>
							</td>

						</tr>
						<tr>
							<td colspan="2" style="border-right:1px solid black;">
								<div class="flexrow">
									<div>Permit Status:</div>
									<div class="value no-u">&nbsp;</div>
								</div>
							</td>
							<td colspan="2">
								<div class="flexrow">
									<div>Issued By:</div>
									<div class="value no-u">&nbsp;</div>
								</div>
							</td>
							
						</tr>
					</table>
					<div class="flexrow">
						<div>Building Use:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Type of Project:</div>
						<div class="value">&nbsp;</div>
					</div>
					<div class="flexrow">
						<div>Improvements:</div>
						<div class="value">&nbsp;</div>
					</div>
					<hr>
					<b>Work Authorized</b>
					<br />
					<table width="100%" class="bdrbtm" style="border-top: 1px solid black;">
						<tr>
							<td align="center" style="border-right:1px solid black;">
								<u>Valuation</u>
								<br />
							</td>
							<td align="center" colspan="2" style="border-right:1px solid black;">
								<u>Construction Type</u>
								<br />
							</td>
							<td align="center">
								<u>Occupancy Group</u>
								<br />
							</td>

						</tr>
					</table>
					<center class="bold">OWNER BUILDER DECLARATION</center>
					<br />
					<input type="checkbox"> I certify that I am the legal owner of the property located at JOB ADDRESS, city of Forest Park, Georgia
					<br />
					<br />
					<input type="checkbox"> I, as the owner of the property am providing my state certification card for performing mechanical, electrical, and plumbing work
					<hr noshade>
					<b>
					This permit shall expire if the building or work authorized by such permit is suspended or abandoned for a period of 180 days.<br />
					<br />
					All work shall conform to the 2012 IBC, IRC, IFC, IPC, IMC, IFCG, 2011 NBC 2009 IECC, 2003 IPMC, 2003 IEBC
					</b>
					<hr noshade>
					<b>Comments/Conditions/Project Description:
					<br />
					<hr noshade>
					<table width="100%">
						<tr>
							<th>Description</th>
							<th>Fee</th>
						</tr>
					</table>
					<br />
					<center class="bold">
						Total Fees: <span style="border: 3px solid black;">$</span>
					</center>
							

				</td>
			</tr>
			<tr>
				<td colspan="2">
					<i>The applicant, his agents and employees of, shale comply with all the rules regulations and requirements of the City Zoning Regulations and Building Codes governing all aspects of the above proposed work for which the permit is granted.  The City of Forest Park or its agents are authorized to order the immediate cessation of construction at anytime a violation of the codes or regulations appears to have occurred.  Violation of any of the codes or regulations applicable may result in the revocation of this permit.  Buildings must conform with submitted and approved plans.  Any changes of plans or layout must be approved prior to changes being made.  Any change in the use or occupancy must be approved prior to commencement of construction.  Construction not commenced within 880 days of permit issuance voids this permit.  Cessation of work for periods of 180 continuous days shall also void this permit.  Permits are not transferable.  The City</i>
				</td>
			</tr>
		</table>
	</body>
</html>
