<style>
	*, body {font-size:16px !important;line-height:26px}
	h1 {font-size:24px}

</style>
<img src="milford_logo.png" width="50%" />
<br />
<center>
	<h1>
	CITY OF MILFORD<br />
	<%=strTITLE%>
	</h1>
	Certificate of Use and Occupancy 4101:1-1-10
</center>
Permit Address: <%
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
<br />
Owner Name: <%=strOwnerName%>
<br />
Owner Address: <%=strOwnerAddress%>
<br />
<br />
This certificate states that the identified building has been constructed, altered or had an addition placed upon it and/or has been inspected and has been found to conform to the applicable provisions of the Ohio Building Code and Chapters 3781. and 3791. Of the Ohio Revised Code.
<br />
<br />
Permit/Plan Approval No.: <%=strPermitNumber%>
<br />
<br />
Building Code: 
<blockquote style="margin:0 40px;">
Edition of the OBC used for review: <%=strBLDQCODE%>
<br />
Use Group(s) 302.0 & Specific Occupancies &mdash; 401.0: <%=oRs("usegroupcode")%>
<br />
Description of Occupancy &mdash; 303.0 to 312.0: <%=strDESCOCC%>
<br />
Construction Type &mdash; Chapter 6: <%=oRs("constructiontype")%>
<br />
Automatic Sprinklers &mdash; 903.0:  <%=strAutoSprinklers%>
<blockquote style="line-height:1px;margin: 0 0 18px 40px;">Hazard Classification: <%=strHazard%></blockquote>
Special Conditions/ Variances Granted:
<br />
<%=strPermitConditions%>
</blockquote>
<br />
Remarks
<br />
<br />
<br />
<br />
		<table class=noborder>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td><%=FormatDateTime(date())%></td>
				<td>&nbsp;</td>
			</tr>
			<tr>
				<td width="60%" style="border-top:1px solid black;">Chief Building Official</td>
				<td>&nbsp;</td>
				<td width="20%" style="border-top:1px solid black;">Date</td>
				<td>&nbsp;</td>
			</tr>
		</table>
