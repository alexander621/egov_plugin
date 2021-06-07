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
		</style>
	</head>
	<%
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<body onLoad="window.print();">
		<table>
			<tr>
				<td colspan="4" class="noborder"><h1>County of Camden</h1><h3>Department of Inspections</h3></td>
				<td align="right" class="noborder">
					PO Box 190<br />
					117 North Highway 343<br />
					Camden, NC 27921<br />
					Phone: 252-338-1919<br />
					Fax: 252-333-1603
				</td>
			</tr>
			<tr>
				<th width="20%">PERMIT NUMBER</th>
				<th width="20%">DATE ISSUED</th>
				<th width="20%">FEE</th>
				<th width="20%">VALUATION</th>
				<th width="20%">ISSUED BY</th>
			</tr>
			<tr>
				<td align="center"><%=strPermitNumber%></td>
				<td align="center"><%=FormatDateTime(oRs("issueddate"),2)%></td>
				<td align="center"><%=FormatCurrency(oRs("FeeTotal"),2)%></td>
				<td align="center"><%=FormatCurrency(oRs("JobValue"),2)%></td>
				<td align="center"><%=strPermitIssuedBy%></td>
			</tr>
		</table>
		<br />
		<table>
			<tr>
				<th rowspan="2" width="2%" style="border-bottom:1px solid black;">N<br />A<br />M<br />E<br />+</th>
				<th rowspan="2" width="2%" style="border-left:none;border-bottom:1px solid black;">A<br />D<br />D<br />R<br />E<br />S<br />S</th>
				<td valign="top" width="48%">LOCATION
					<br />
					<%
					response.write oRs("ListedOwner")
					response.write "<br />"
					response.write "<br />"
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
				<td style="border-left:none" width="48%">
					<table class="noborder">
						<tr>
							<td align="right" width="25%">PIN</td>
							<td colspan="3"><%=oRs("parcelidnumber")%></td>
						</tr>
						<tr>
							<td align="right">SUBDIVISION</td>
							<td colspan="3"><%=oRs("county")%></td>
						</tr>
						<tr>
							<td align="right">LOT #</td>
							<td colspan="3"><%=strLotnumber%></td>
						</tr>
						<tr>
							<td align="right"><nobr>PROPERTY SIZE</nobr></td>
							<td><%=strPropertysize%>Sq. Ft.</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td valign="top">APPLICANT
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
				</td>
				<td valign="top">CONTRACTOR
					<br />
					<%
						if oRs("isprimarycontractor") then
							If oRs("confirstname") <> "" Then 
								response.write oRs("confirstname") & " " & oRs("conlastname") & "<br />"
							End If 
							If Not IsNull(oRs("concompany")) And oRs("concompany") <> "" Then 
								response.write oRs("concompany") & "<br />"
							End If 
							If Not IsNull(oRs("conaddress")) And oRs("conaddress") <> "" Then 
								response.write  oRs("conaddress") & "<br />"
							End If 
							If Not IsNull(oRs("concity")) And oRs("concity") <> "" Then
								response.write  oRs("concity") & ", " & oRs("constate") & " " & oRs("conzip") & "<br />" 
							End If 
							response.write FormatPhoneNumber(oRs("conphone"))
						else
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
							response.write FormatPhoneNumber(oRs("appphone"))
						end if
					%>
				<br />
				<br />
				LICENSE # <%=GetPrimaryContactLicense(intPermitID)%>
				</td>
			</tr>
		</table>
		<br />
		<table>
			<tr><th colspan="2">SELECTED CHARACTERISTICS OF WORK</th></tr>
			<tr>
				<td valign="top" width="50%">
					PERMIT TYPE: <%=oRs("permittypedesc")%><br />
					<%=oRs("DescriptionOfWork")%><br />
					UDO NUMBER:<br />
					<%=strUdonumber%><br />
					FLOOD ZONE:<br />
					<%=strFloodzone%><br />
					FEMA requirement: lowest utility should be<br />
					Base Flood Elevation plus 1 foot<br />
					ZONING:<br />
					<%=oRs("Zoning")%><br />
					<table class="noborder">
						<tr>
							<td rowspan="2" valign="top">SETBACKS:</td>
							<td align="center">FRONT</td>
							<td align="center">SIDE</td>
							<td align="center">REAR</td>
						</tr>
						<tr>
							<td align="center"><%=strFrontsetback%></td>
							<td align="center"><%=strSidesetback%></td>
							<td align="center"><%=strRearsetback%></td>
						</tr>
					</table>
				</td>
				<td valign="top" width="50%">
					<center>SUB-CONTRACTORS</center>
					<table class="noborder">
						<tr><td width="80%">ELECTRICAL</td><td width="20%">License #</td></tr>
						<tr><td><%=strEleConName%></td><td><%=strEleConLic%></td></tr>
						<tr><td width="80%">MECHANICAL</td><td width="20%">License #</td></tr>
						<tr><td><%=strMechConName%></td><td><%=strMechConLic%></td></tr>
						<tr><td width="80%">PLUMBING</td><td width="20%">License #</td></tr>
						<tr><td><%=strPlumbConName%></td><td><%=strPlumbConLic%></td></tr>
						<tr><td width="80%">INSULATION</td><td width="20%">License #</td></tr>
						<tr><td><%=strInsuConName%></td><td><%=strInsuConLic%></td></tr>
					</table>
					TOTAL FLOOR AREA OF NEW CONST. <%=FormatNumber(oRs("totalsqft"),2)%> Sq. Ft.
				</td>
			</tr>
		</table>
		<br />
		<table>
			<tr><th>AFFADAVIT OF APPLICANT</th></tr>
			<tr>
				<td>
				The following items shall be required before final inspection and certificate of occupancy is released.<br />
				<ol>
					<li>Certificate of Elevation (if required)</li>
					<li>Certificate of Authorized Contractors (Electrical, Mechanical, Plumbing) attached as "Exhibit A" and by references incorporated herein as if set forth verbatim.</li>
				</ol>
				I hereby certify that I have the authority to make the necessary applications, that the applications are correct, and that the construction will conform to the regulations in the Building, Electrical, Plumbing, and Mechanical Codes, and all other LOCAL and STATE laws and/or ordinances.
				<br />
				I do certify and guarantee that prior to the commencement of any work performed, I/we shall have obtained the necessary permits authorizing said work and also do acknowledge that unless I fully comply with all STATE and LOCAL regulations that relate to Electrical, Plumbing, Insulation, and Mechanical codes and all appurtenant laws and regulations governing those permits heretofore issued; they shall be void and of no further force or effect. This shall result in the automatic revocation of all permits or authorizations issued. If I/we have not obtained the necessary permits required prior to commencement of any work performed this too shall result in automatic revocation of any permits and all authorizations to proceed with work.
				<br />
				<br />
				<br />
				<table class="noborder">
					<tr>
						<td></td>
						<td width="1%"></td>
						<td></td>
						<td width="1%"></td>
						<td align="center" width="10%"><%=FormatDateTime(oRs("issueddate"),2)%></td>
					</tr>
					<tr>
						<td style="border-top: 1px solid black;margin:20px;">Issuing Officer / Permit Clerk</td>
						<td></td>
						<td style="border-top: 1px solid black;margin:20px;">Signature of Applicant</td>
						<td></td>
						<td style="border-top: 1px solid black;margin:20px;">Date</td>
					</tr>
				</table>
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
