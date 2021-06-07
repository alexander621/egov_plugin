<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<style>
			body
			{
				font-family: Ariel, Ariel, sans-serif;
				-webkit-print-color-adjust:exact;
			}
			table
			{
				width:100%;
				border-spacing: 0;
			}
			h1,h2,h3,h4
			{
				margin:0;
			}
		</style>
	</head>
	<%
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<body onLoad="window.print();">
		<table>
			<tr>
				<td align="center"><img src="logo.png" height="125" /></td>
				<td align="center" valign="top" width="50%">
				<br />
				<br />
				<br />
					<h2><%=strTitle%>TEMPORARY CERTIFICATE OF OCCUPANCY</h2>
				</td>
				<td align="center">
					City of Wyoming<br />
					800 Oak Avenue<br />
					Wyoming, OH 45215<br />
					(513) 821-7600<br />
					www.wyomingohio.gov<br />
				</td>
			</tr>
		</table>
		<br />
		<br />
		<br />
		<table style="border:1px solid black;">
			<tr>
				<td width="50%">Permit Number: <%=strPermitNumber%></td>
				<td width="50%">Date Permit Issued: <%=FormatDateTime(oRs("issueddate"),2)%></td>
			</tr>
			<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
			<tr>
				<td width="50%" valign="top">
					<table>
						<td valign="top">Project Address:</td>
						<td>
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
					response.write "<br />"
					
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
						</td>
					</table>
				</td>
				<td width="50%" valign="top">
					<table>
						<td valign="top">Owner:</td>
						<td>
						<%=strOwnerAddress%>
						</td>
					</table>
				</td>
			</tr>
			<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
			<tr>
				<td>Approved As: <%=oRs("approvedas")%></td>
			</tr>
			<tr><td>&nbsp;</td><td>&nbsp;</td></tr>
			<tr>
				<td width="50%" valign="top">Description of Work: <%=oRs("descriptionofwork")%></td>
				<td width="50%" valign="top">Date T.C.O. Issued: <%=Date()%></td>
			</tr>
		</table>
		<br />
		<br />
		<br />
		<table style="border:1px solid black;">
			<tr>
				<td>
					Stipulations, Conditions, Variances: 
					<br />
					<%
					tempconotes = oRs("tempconotes") & ""
					%>
					<%=replace(tempconotes,vbcrlf,"<br />")%>
					<br />
				</td>
			</tr>
		</table>
		<br />
		<br />
		<br />
		<table style="border:1px solid black;">
			<tr>
				<td>
				This Certificate represents an approval that is valid only when the building and its facilities are used as stated and is conditional upon all building systems being maintained and tested in accordance with the applicable Ohio Board of Building Standards rules and applicable equipment or system schedules.
				</td>
			</tr>
		</table>
		<br />
		<br />
		<br />
		Approved pursuant to the following codes and/or regulations:<br />
		<input type="checkbox" <%if blnApprovedOhio then response.write " checked"%> />Residential Code of Ohio &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" <%if blnApprovedWyoming then response.write " checked"%> />Wyoming Zoning Code
		<br />
		<br />
		<br />
		<br />
		<br />
		<br />
		_______________________________________________________________<br />
		Daniel Bly, Residential Building Official<br />
		Megan Statt Blake, Community Development Director
		<br />
		<br />
		<br />
		<center>
		<h4>City of Wyoming, Ohio</h4>
		<hr style="height:3px;border:none;color:#333;background-color:#333;margin:0;" />
		<h3>Community Development Department • Certified Residential Building Department</h3>
		</center>
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
