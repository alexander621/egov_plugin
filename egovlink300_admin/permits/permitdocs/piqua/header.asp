		<table>
			<tr>
				<td width="50%"><img src="piqua-logo.png" width="80%" border="0" /></td>
				<td align="right" valign="center" width="50%"><h3><%=strTitle%></h3></td>
			</tr>
			<tr>
				<th>LOCATION INFORMATION</th>
				<th>PERMIT</th>
			</tr>
			<tr>
				<td>
					Street Address: 
					
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
					
					If oRs("residentcity") <> "" Then
						response.write ", " & oRs("residentcity")
					End If 
					If oRs("residentstate") <> "" Then
						response.write ", " & oRs("residentstate")
					End If 
					If oRs("residentzip") <> "" Then 
						response.write " " & oRs("residentzip")
					End If 
					%>
				</td>
				<td>Number: <%=strPermitNumber%></td>
			</tr>
			<tr>
				<td>Owner Name: <%=oRs("listedowner")%></td>
				<td>Status: <%=oRs("permitstatus")%></td>
			</tr>
			<tr>
				<td>Parcel ID: <%=oRs("parcelidnumber")%></td>
				<td>Date: <%=date()%></td>
			</tr>
		</table>
		<br />
