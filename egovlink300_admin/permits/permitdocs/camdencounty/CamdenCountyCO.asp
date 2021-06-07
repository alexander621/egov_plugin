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
	<body onload="window.print();">
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
	CERTIFICATE OF OCCUPANCY<br />
NORTH CAROLINA GENERAL STATUTE 160A-423</h2></center>
<br />
<br />
<br />
		Date of Issue:&nbsp;&nbsp;&nbsp; <%if not isnull(oRs("completeddate")) then response.write FormatDateTime(oRs("completeddate"),2) else response.write "PERMIT NOT COMPLETE" end if%>
		<br />
		<br />
		Owner:&nbsp;&nbsp;&nbsp; <%=oRs("listedowner")%>
		<br />
		<br />
		Type of Occupancy:&nbsp;&nbsp;&nbsp; <!--<%=oRs("usegroupcode")%>&nbsp;<%=oRs("occupancytype")%>--> <%=oRs("permittypedesc")%>
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
		Building Permit #:&nbsp;&nbsp;&nbsp; <%=strPermitNumber%>
		<br />
		<br />
		<hr noshade />
		<br />
		<br />
		Signature:  _________________________________________________________________
		<br />
		<br />
Building Code Enforcement Officer<br />
PO Box 190<br />
117 North Highway 343<br />
Camden, NC 27921<br />
Phone: 252-338-1919<br />
Fax: 252-333-1603<br /><br />
To the best of my knowledge and belief, said building was in substantial compliance with the North Carolina State Codes and/or Camden County Code at the time of the final inspection
<br /><br />
The subject construction was not inspected for compliance with the Americans with Disabilities Act (ADA). Compliance with the said Act is the sole responsibility of the property owner.
	</body>
</html>
