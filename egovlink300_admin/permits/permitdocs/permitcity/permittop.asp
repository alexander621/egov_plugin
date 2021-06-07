
		<table class="noborder">
			<tr>
				<td class="noborder" width="50%"><img src="milford_logo.png" width="100%" /></td>
				<td align="right" class="noborder"><h2><%=strTITLE%></h2></td>
			</tr>
		</table>
		<table>
			<% if strTITLE = "PLAN APPROVAL" then%>
			<tr>
				<td colspan="4" align="center"><b>CALL BUILDING DEPARTMENT 513-248-5098 FOR INSPECTION</b></td>
			</tr>
			<% end if %>
			<tr>
				<td colspan="2"><b>LOCATION INFORMATION</b></td>
				<td colspan="2"><b>PERMIT</b></td>
			</tr>
			<tr>
				<td>Street Address:</td>
				<td class="noleftborder">
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
				</td>
				<td rowspan="3" class=" borderbottom">Number:</td>
				<td rowspan="3" class="noleftborder borderbottom"><%=strPermitNumber%></td>
			</tr>
			<tr>
				<td>Parcel ID:</td>
				<td class="noleftborder norightborder"><%=oRs("parcelidnumber")%></td>
			</tr>
			<tr>
				<td>Zoning:</td>
				<td class="noleftborder norightborder"><%=oRs("zoning")%></td>
			</tr>
			<% if strTITLE = "ZONING CERTIFICATE" then%>
			<tr>
				<td>Existing Use:</td>
				<td class="noleftborder"><%=oRs("existinguse")%></td>
				<td style="border-top:0;">Proposed Use:</td>
				<td class="noleftborder" style="border-top:0;"><%=oRs("proposeduse")%></td>
			</tr>
			<% end if %>
		</table>
		<br />
		<table>
			<tr>
				<td width="50%"><b>APPLICANT INFORMATION</b></td>
				<td width="50%"><b>PROPERTY OWNER INFORMATION</b></td>
			</tr>
			<tr>
				<td valign="top">
					<br />
					Name:<%=oRs("appfirstname") & " " & oRs("applastname")%><br />
					<br />
					Address:<%=oRs("appaddress") & " " & oRs("appcity") & ", " & oRs("appstate") & " " & oRs("appzip") %><br />
					<br />
					Phone:<%=oRs("appphone")%><br />
					<br />
					Email:<%=oRs("appemail")%><br />
					<br />

				</td>
				<td valign="top">
					<br />
					Name:<%=strOwnerName%><br />
					<br />
					Address:<%=strOwnerAddress %><br />
					<br />
					Phone: <%=strOwnerPhone%><br />
					<br />
					Email: <%=strOwnerEmail%><br />
					<br />
				</td>
			</tr>
		</table>
