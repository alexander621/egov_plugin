<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<style>
			body
			{
				font-family: Ariel, Ariel, sans-serif;
				-webkit-print-color-adjust:exact;
				font-size: 12px !important;
				margin:0px;
			}
			table
			{
				font-size: 12px !important;
				width:100%;
				border-spacing: 0;
			}
			th {
				background-color:#ccc;
			}
			@media print {
				body
				{
					color-adjust: exact;
				}
			th {
				background-color:#ccc;
			}

			}
			h2 {font-weight:normal;}
			h1,h2,h3
			{
				margin:0;
			}
			td, th
			{
				border-top: 1px solid black;
				border-left: 1px solid black;
				padding: 3px;
			}
			td:last-child, th:last-child
			{
				border-right: 1px solid black;
			}
			tr:last-child td
			{
				border-bottom: 1px solid black;
			}
			td.norightborder
			{
				border-right:0 !important;
			}
			td.borderbottom
			{
				border-bottom: 1px solid black;
			}
			table.noborder td, table.noborder th, td.noborder, th.noborder
			{
				border:none;
			}
			hr
			{
				color: #FFF;
				background-color: #FFF;
				border:0;
				page-break-after: always;
			}

			.noleftborder { border-left:none; }
			.notopborder { border-top:none; }
		</style>
	</head>
	<%
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<body onLoad="window.print();">
		<% strTITLE="PLAN APPROVAL" %>
		<!--#include file="permittop.asp"-->
		<br />
		<table>
			<tr>
				<td colspan="2"><b>PROJECT INFORMATION</b></td>
			</tr>
			<tr>
				<td colspan="2">Description of Work: <%=oRs("descriptionofwork")%></td>
			</tr>
			<tr>
				<td>Construction Type: <%=oRs("constructiontype")%></td>
				<td>Occupancy Type/Use Gorup: <%=oRs("usegroupcode")%> <%=oRs("occupancytype")%>/<%=oRs("usegroupcode")%></td>
			</tr>
			<tr>
				<td>Work Scope: <%=oRs("workscope")%></td>
				<td>Work Class: <%=oRs("workclass")%></td>
			</tr>
			<tr>
				<td>Estimated Cost: <%=oRs("jobvalue")%></td>
				<td>Occupant Load/Dwelling Units: <%=oRs("occupants")%>/<%=oRs("residentialunits")%></td>
			</tr>
			<tr>
				<td>Automatic Sprinklers: <%=strAutoSprinklers%></td>
				<td>Hazard Classification: <%=strHazard%></td>
			</tr>
			<tr>
				<td colspan="2">
					Building is to be <%=oRs("structurewidth")%> Ft. Wide By <%=oRs("structurelength")%> Ft. Long By <%=oRs("structureheight")%>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					Area or Volume (Cubic/Square Feet): <%=oRs("totalsqft")%> sq.ft.
				</td>
			</tr>

		</table>
		<br />
		<table>
			<tr>
				<td><b>FEES</b></td>
			</tr>
		</table>
		<table class="noborder" style="border:1px solid black;border-top:0;">
			<tr>
				<td align="right" width="20%">PERMIT = </td>
				<td><%=dPermitFee%></td>
				<td width="20%"></td>
			</tr>
			<tr>
				<td align="right">C.O. = </td>
				<td><%=dCofOFee%></td>
				<td></td>
			</tr>
			<tr>
				<td align="right">OBC/RCO = </td>
				<td><%=dBBSFee%></td>
				<td></td>
			</tr>
			<tr>
				<td align="right">PARKS & REC FEE = </td>
				<td><%=dRecImpactFee%></td>
				<td></td>
			</tr>
			<tr>
				<td></td>
				<td align="right">PERMIT FEES TOTAL = </td>
				<td><%=dNoZoneTotalFees%></td>
			</tr>
		</table>
		<br />
		<table>
			<tr>
				<td><b>ADDITIONAL COMMENTS:</b></td>
			</tr>
			<tr>
				<td>
				<b style="color:red;">***Plumbing Inspection Approvals: Clermont Co. Board of Health, 513-732-7499<br />
				***Electrical Inspection Approvals: IBI, 513-381-6080	</b>
				<br />
				<%=replace(oRs("permitnotes"),vbcrlf,"<br />")%>
				<br />
				<br />
				<br />
				</td>
			</tr>
		</table>
		This permit is granted on the express condition that the said construction shall in all respects, conform to the Ordinances of this jurisdiction including the Zoning Ordinance, regulating the construction and use of buildings, and may be revoked at any time upon violation of any provisions of said ordinances. The Department reserves the right to reject any work which has been concealed or completed without first having been inspected and approved by the Department. Any deviation from the approved plans must be authorized by the approval of revised plans subject to the same procedure established for the examination of the original plans. An additional permit fee is also charged predicated on the extent of the variation from the original plans. Permits are not valid if construction work is not started within six months from date permit is issued. Final Inspection and certificate of occupancy must be obtained before occupying the building.
		<br />
		<br />
		<table>
			<tr>
				<td width="80%">Chief Building Official:</td>
				<td>
					Date: <% if not isnull(oRs("issueddate")) then response.write FormatDateTime(oRs("issueddate"))%>
				</td>
			</tr>
		</table>
		




		<% 
			oRs.Close
			Set oRs = Nothing
			oRsA.Close
			Set oRsA = Nothing
			oRsPR.Close
			Set oRsPR = Nothing
		%>
	</body>
</html>
