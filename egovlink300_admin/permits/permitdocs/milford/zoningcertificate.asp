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
		<% strTITLE="ZONING CERTIFICATE" %>
		<!--#include file="permittop.asp"-->
		<br />
		<table>
			<tr>
				<td><b>PROJECT INFORMATION</b></td>
			</tr>
			<tr>
				<td>
					Description of Work:
					<br />
					<%=oRs("descriptionofwork")%>
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
				<td align="right" width="20%"><nobr>ZONING FEES TOTAL = </nobr></td>
				<td><%=formatcurrency(dZoneFeeAmount)%></td>
			</tr>
		</table>
		<br />
		<table>
			<tr>
				<td><b>ADDITIONAL COMMENTS:</b></td>
			</tr>
			<tr>
				<td>
				<%=replace(strZAddComments,vbcrlf,"<br />")%>
				<br />
				<br />
				<br />
				</td>
			</tr>
		</table>
		I have examined the foregoing application, plans, materials, and specifications and hereby approve them for compliance to the Planning and Zoning Code of the City of Milford.
		<br />
		<br />
		<table>
			<tr>
				<td width="80%">Zoning Administrator:</td>
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
