<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<style>
			body
			{
				font-family: Ariel, Ariel, sans-serif;
				-webkit-print-color-adjust:exact;
				font-size: 16px !important;
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
			h1,h2,h3
			{
				margin:0;
			}
			td, th
			{
				border-top: 1px solid black;
				border-left: 1px solid black;
				padding: 5px;
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
		</style>
	</head>
	<%
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<!--body onload="window.print();"-->
	<body>
	<table class="noborder">
		<tr>
			<td width="205"><img src="camden-logo.png" width="200" /></td>
			<td>
			<h1>County of Camden, NC</h1>
			<h2>Department of Inspections</h2>
			</td>
		</tr>
	</table>
	</table>
	<center><h2>
	FLOOD PLAIN CERTIFICATE OF COMPLIANCE</h2></center>
	<br />
	<br />
This certificate is issued pursuant to the requirements of the North Carolina State Building Code and G.S. 153A-363 for the following:
<br />
<br />
<br />


Location:&nbsp;&nbsp;&nbsp;
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
					 response.write ", "
					
					If oRs("residentcity") <> "" Then
						response.write oRs("residentcity")
					End If 
					If oRs("residentstate") <> "" Then
						response.write ", " & oRs("residentstate")
					End If 
					If oRs("residentzip") <> "" Then 
						response.write " " & oRs("residentzip")
					End If 
					%>
<br />
<br />

Permit Number:&nbsp;&nbsp;&nbsp;<%=strPermitNumber %>
<br />
<br />

Property Owner:	<%=oRs("listedowner")%>
<br />
<br />
<br />
<br />
<br />
<br />




Type of Occupancy:&nbsp;&nbsp;&nbsp;<%=oRs("permittypedesc")%>
<br />
<br />
<br />
<br />
<br />
<br />
<br />



<table style="width: 50%; font-size:16px !important;">
<tr>
	<td colspan="2" style="border:0;"></td>
	<td style="border:0;"><%=date()%></td>
</tr>
<tr>
	<td style="border:0;border-top:1px solid black;">Building Inspector<br />Camden County</td>
	<td style="border:0;">&nbsp;&nbsp;</td>
	<td style="border:0;border-top:1px solid black;">Date</td>
</tr>
</table>
<br />
<br />
Occupancy Max required for non residential structures only.<br />
Violation of this certificate of occupancy constitutes a Class 1 misdemeanor.


	</body>
</html>
