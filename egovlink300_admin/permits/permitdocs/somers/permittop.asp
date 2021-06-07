
		<table class="noborder">
			<tr>
				<td class="noborder" width="50%"><img src="logo.png" width="100%" /></td>
				<td align="right" class="noborder"><h2><%=strTITLE%></h2></td>
			</tr>
		</table>
		<table>
			<tr>
				<td colspan="2"><b>LOCATION INFORMATION</b></td>
				<td colspan="2"><b>PERMIT</b>&nbsp;&nbsp;&nbsp;<%=oRs("permittypedesc")%></td>
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
				<td>Number:</td>
				<td class="noleftborder"><%=strPermitNumber%></td>
			</tr>
			<tr>
				<td>Owner Name:</td>
				<td class="noleftborder"><%=oRs("ListedOwner")%></td>
				<td>Status:</td>
				<td class="noleftborder"><%=oRs("permitstatus")%></td>
			</tr>
			<tr>
				<td>Parcel ID:</td>
				<td class="noleftborder"><%=oRs("parcelidnumber")%></td>
				<td>Date:</td>
				<td class="noleftborder"><%=FormatDateTime(oRs("issueddate"),2)%></td>
			</tr>
		</table>
		<table>
			<tr>
				<td width="50%" class="notopborder"><b>APPLICANT INFORMATION</b></td>
				<td width="50%" class="notopborder"><b>OWNER INFORMATION</b></td>
			</tr>
			<tr>
				<td valign="top">
					<br />
					<%
						If oRs("appfirstname") <> "" Then 
							response.write oRs("appfirstname") & " " & oRs("applastname") & "<br />"
						End If 
						If Not IsNull(oRs("appcompany")) And oRs("appcompany") <> "" Then 
							response.write oRs("appcompany") & "<br />"
						End If 
						If Not IsNull(oRs("appaddress")) And oRs("appaddress") <> "" Then 
							response.write  oRs("appaddress") & "<br />"
						End If 
						If Not IsNull(oRs("appcity")) And oRs("appcity") <> "" Then
							response.write  oRs("appcity") & ", " & oRs("appstate") & " " & oRs("appzip") & "<br />" 
						End If 
					%>
					<br />
					<br />
					<br />

				</td>
				<td valign="top"><br /><%=oRs("ListedOwner")%></td>
			</tr>
