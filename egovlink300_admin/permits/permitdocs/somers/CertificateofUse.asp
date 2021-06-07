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
		<% strTITLE="CERTIFICATE OF USE" %>
		<!--#include file="permittop.asp"-->
			<tr>
				<td><b>PROJECT INFORMATION</b></td>
				<td></td>
			</tr>
			<tr>
				<td>Type of Construction: <%=oRs("constructiontype")%></td>
				<td>Description of Work: <%=oRs("descriptionofwork")%></td>
			</tr>
			<tr>
				<td>Occupancy Use: <%=oRs("occupancytype")%></td>
				<td>Occupants: <%=oRs("occupants")%></td>
			</tr>
		</table>
		<br />
		<table>
			<tr>
				<td>
					Stipulations, Conditions, Variances: <br />
					<%=oRs("conotes")%>
					<br />
					<br />
					<br />
					<br />
					<br />
					<br />
				</td>
			</tr>
		</table>
		<table>
		<br />
		<br />
		<br />
		<br />
		The changes or work authorized by the permit issued for this project have been inspected and approved as noted above. The issuance of this Certificate of Use acknowledges the completeness and appropriateness of the changes and or work authorized by the above referenced permit and with reference to any conditions noted above.
		<br />
		<br />
		<br />
		<br />
		Approving Official ___________________________________________________ Date Issued ___________________________
		




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
