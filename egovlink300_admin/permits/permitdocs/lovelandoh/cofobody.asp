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
	<body onload="window.print();">
		<table>
			<tr>
				<td align="right"><img src="loveland-logo-2.png" height="125" /></td>
				<td align="center" valign="top" width="50%">
					<h2><%=strTitle%>Certificate of Occupancy</h2>
					<h3>Office of the Building Official</h3>
				</td>
				<td>
				<b>City of Loveland<br />
				Building & Zoning</b><br />
				120 W. Loveland Ave.<br />
				Loveland, Ohio 45140<br />
				Office: 513.707.1447<br />
				Fax: 513.583.3032<br />
				www.lovelandoh.com
				</td>
			</tr>
		</table>
		<br />
		<br />
		<br />
		<table style="border:1px solid black;">
			<tr>
				<td width="50%">Permit Number: <%=strPermitNumber%></td>
				<td width="50%">Date Issued: 
					<% if instr(strTitle, "Temporary") > 0 then %>
						<%=FormatDateTIme(date(),2)%>
					<%elseif strFinalInspectionDate <> "" then %>
						<%=FormatDateTime(strFinalInspectionDate,2)%>
					<% end if %>
				</td>
			</tr>
			<tr>
				<td>Approved As: <%=oRs("approvedas")%></td>
				<td>Occupancy Use: <%=oRs("usegroupcode")%>&nbsp;<%=oRs("occupancytype")%></td>
			</tr>
			<tr>
				<td>Type of Construction: <%=oRs("constructiontype")%></td>
				<td>Occupants: <%=oRs("occupants")%></td>
			</tr>
		</table>
		<br />
		<br />
		<table>
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
				<%
					response.write oRs("ListedOwner")
					'PDF ONLY INCLUDED 'owner' which translates to the "Listed Owner" field
					'response.write "<br />"
					'response.write oRs("residentstreetnumber")
					'If oRs("residentstreetprefix") <> "" Then
						'response.write " " & oRs("residentstreetprefix")
					'End If
					'response.write " " & oRs("residentstreetname")
					'If oRs("streetsuffix") <> "" Then
						'response.write " " & oRs("streetsuffix")
					'End If
					'If oRs("streetdirection") <> "" Then
						'response.write " " & oRs("streetdirection")
					'End If
					'If oRs("residentunit") <> "" Then
						'response.write ", " & oRs("residentunit")
					'End If
					'response.write "<br />"
					'
					'If oRs("residentcity") <> "" Then
						'response.write oRs("residentcity")
					'End If 
					'If oRs("residentstate") <> "" Then
						'response.write ", " & oRs("residentstate")
					'End If 
					'If oRs("residentzip") <> "" Then 
						'response.write " " & oRs("residentzip")
					'End If 
					'response.write "<br />"
						'response.write FormatPhoneNumber(oRs("appphone"))
					%>
						</td>
					</table>
				</td>
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
					<% if strTitle = "Temporary " then %>
					<%=oRs("tempconotes")%>
					<% else %>
					<%=oRs("conotes")%>
					<% end if %>
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
		Approved pursuant to the following editions of:<br />
		<u><%=strEdition%></u>
		<br />
		<br />
		<br />
		<br />
		<br />
		<br />
		_______________________________________________________________<br />
		James McFarland, Chief Building Official
		<br />
		<br />
		<br />
		<center>
		<h4>State of Ohio &ndash; City of Loveland</h4>
		<hr style="height:3px;border:none;color:#333;background-color:#333;margin:0;" />
		<h3>Certified Building Department</h3>
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
