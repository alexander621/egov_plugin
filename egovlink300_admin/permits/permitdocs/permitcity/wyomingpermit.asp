<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<style>
			body
			{
				font-family: Ariel, Ariel, sans-serif;
				font-size: 12px;
				-webkit-print-color-adjust:exact;
			}
			table
			{
				width:100%;
				border-spacing: 0;
				margin-bottom:5px;
			}
			h1,h2,h3,h4
			{
				margin:0;
			}
			td, th
			{
				border-top: 1px solid black;
				border-left: 1px solid black;
				padding: 7px;
				font-size:10px;
			}
			td b, th b
			{
				font-size:12px;
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
		<table style="border-bottom:1px solid black;">
			<tr>
				<td align="right" class="noborder" width="10%"><img src="logo.png" height="75" /></td>
				<td align="center" class="noborder" valign="top" width="80%">
					<h3 style="font-size:14px;">
					<br />
						RESIDENTIAL BUILDING PERMIT <br />
						PUBLIC AREA EXCAVATION PERMIT<br />
						ZONING CERTIFICATE
					</h3>
				</td>
				<td align="center" width="10%" class="noborder">
					City of Wyoming<br />
					800 Oak Avenue<br />
					Wyoming, OH 45215<br />
					(513) 821-7600<br />
					www.wyomingohio.gov<br />
				</td>
			</tr>
		</table>
		<br />
		<center>
			<b>PERMIT NUMBER: <u><%=strPermitNumber%></u></b>
		<br />
		<u><b style="font-size:11px;">** APPROVED PLANS ARE REQUIRED TO BE ON SITE FOR EVERY INSPECTION ** FOR INSPECTIONS PLEASE REFER TO PAGE 2</b></u>
		</center>
		<br />
		<table>
			<tr>
				<td style="border-right:solid 1px black;" width="49%">
					<h4 style="display:inline-block;">JOB SITE ADDRESS:</h4>
					<%=oRs("residentstreetnumber")%>
					<%If oRs("residentstreetprefix") <> "" Then
						response.write " " & oRs("residentstreetprefix")
					End If
					response.write " " & oRs("residentstreetname")
					If oRs("streetsuffix") <> "" Then
						response.write " " & oRs("streetsuffix")
					End If
					If oRs("streetdirection") <> "" Then
						response.write " " & oRs("streetdirection")
					End If%>
				</td>
				<td class="noborder" width="2%"></td>
				<td class="noborder" width="49%"></td></tr>
					
			<tr>
				<td style="border-right:solid 1px black;" valign="top">
					<b>Property Owner:</b><br />
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
				<td class="noborder" style="border-bottom:0;"></td>
				<td valign="top">
					<b>Primary Contractor:</b><br />
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
				</td>
			</tr>

		</table>
		<br />
		<table>
			<tr>
				<td width="49%" style="border-right:solid 1px black;" valign="top">
					<b>Primary Contact:</b><br />
					<br />
					<%=oRs("primarycontact")%>
				</td>
				<td width="2%" class="noborder" style="border-bottom:0;"></td>
				<td width="49%" valign="top">
					<b>Plans Prepared By:</b><br />
					<br />
					<%=sPlansBy%>
				</td>
			</tr>

		</table>
		<table class="noborder">
			<tr>
				<td align="right"> Code under which plans were approved: </td>
				<td><input type="checkbox" <%if blnApprovedOhio then response.write " checked"%> />Residential Code of Ohio &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" <%if blnApprovedWyoming then response.write " checked"%> />Wyoming Zoning Code</td>
			</tr>
			<tr>
				<td align="right">Other:</td>
				<td style="border-bottom:1px solid black;"><%=strPlansApprovedOther%>&nbsp;</td>
			</tr>
		</table>
		<table>
			<tr>
				<td>
					<b>Description of proposed work:</b><br />
					<br />
					<%=oRs("descriptionofwork")%>
				</td>
			</tr>
		</table>
		<table class="noborder">
			<tr>
				<td>Project Valuation: <%=formatcurrency(oRs("jobvalue"))%></td>
				<td>Permit Fee: <%=dPermitFee%></td>
			</tr>
		</table>
		<table>
			<tr>
				<td>
					<b>Conditions of Permit Approval:</b><br />
					<br />
					<%=strPermitConditions%>&nbsp;
				</td>
			</tr>
		</table>
		<b>ALL WORK SHALL BE IN COMPLIANCE WITH APPLICABLE BUILDING & ZONING REGULATIONS, THE ORDINANCES, AND STANDARDS OF THE CITY OF WYOMING DEPARTMENTS OF PUBLIC WORKS & WATER WORKS AS APPLICABLE, AND AS APPROVED BY THE CITY OF WYOMING COMMUNITY DEVELOPMENT DEPARTMENT, A STATE OF OHIO CERTIFIED RESIDENTIAL BUILDING DEPARTMENT.</b>
		<br />
		The approval of plans, drawings, and/or specifications in accordance with this permit is invalid if construction, erection, alteration, or other work upon the building has not commenced within twelve months of the approval of the plans, drawings, and/or specifications. One extension shall be granted for an additional twelve-month period if requested by the owner at least ten days in advance of the expiration of the approval and upon payment of a fee not to exceed one hundred dollars. If in the course of construction, work is delayed or suspended for more than six months, the approval of plans or drawings and specifications or data is invalid. Two extensions shall be granted for six months each if requested by the owner at least ten days in advance of the expiration of the approval and upon payment of a fee for each extension of not more than one hundred dollars.
		<br />
		<br />
		<table class="noborder">
			<tr>
				<td>Approving Official: _______________________________________</td>
				<%
				issueddate = ""
				if not isnull(oRs("issueddate")) then issueddate = formatdatetime(oRs("issueddate"),2)
				%>
				<td>Date Issued: <u><%=issueddate%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></td>
			</tr>
		</table>
		<hr>
		<center>
				<b style="font-size:14px;">INSPECTION INFORMATION FOR HOLDERS OF<br />
				RESIDENTIAL BUILDING PERMITS, PUBLIC AREA EXCAVATION PERMITS & ZONING CERTIFICATES</b>
		</center>
		<br />
		<br />
		Certain inspections are required to be conducted during the construction process to ensure compliance with the applicable provisions of the Building and Zoning Codes and/or the approved plans. The following is a list of the specific inspections required for your project:
		<table>
			<tr>
				<td>
					<b>PROJECT TYPE:</b><%=oRs("permittypedesc")%><br />
					<b>REQUIRED INSPECTIONS:</b><br />
					<br />
			<%
				Do While Not oRsI.EOF
				%>
					<%=oRsI("inspection")%><br /><br />
					<%oRsI.MoveNext
				loop
				oRsI.Close
				Set oRsI = Nothing
			%>
				</td>
			</tr>
		</table>
		<center style="font-size:16px;">
			INSPECTION REQUEST NUMBERS: <br />
			For Building Permits with a "B" Prefix: (513) 842-1398<br />
			For Zoning Certificates with a "Z" Prefix: (513) 821-7600<br />
		</center>
Your request must be scheduled a minimum of 24 hours prior to the desired time of the inspection, excluding weekends and legal holidays. Inspection requests left after hours may not be performed until the second business day after the message was left depending on inspector availability.
<br />
<br />
<u>When scheduling an inspection, please provide the following information or leave it in your message:</u>
<ul>
<li>Your name and company name (if applicable).</li>
<li>The phone number where you can be reached during regular business hours.</li>
<li>What type of inspection you are requesting (see above).</li>
<li>Address of the job: <u>
					<%=oRs("residentstreetnumber")%>
					<%If oRs("residentstreetprefix") <> "" Then
						response.write " " & oRs("residentstreetprefix")
					End If
					response.write " " & oRs("residentstreetname")
					If oRs("streetsuffix") <> "" Then
						response.write " " & oRs("streetsuffix")
					End If
					If oRs("streetdirection") <> "" Then
						response.write " " & oRs("streetdirection")
					End If%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
 Your permit number: <u><%=strPermitNumber%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</u></li>
<li>The desired date of the requested inspection *</li>
<li>Location of any blueprints and/or drawings on the property (if applicable) and any special instructions you need the inspector to know.</li>
</ul>
<i>*Inspections are generally performed in the afternoon. Depending on the inspector’s work load on any given day, inspections may not be conducted until early evening. The inspector will generally not call you to confirm the inspection or to schedule a specific time. If you need to request a specific time or would like a return call from the inspector, please specify that in your message.</i>
<br />
<br />
<center><u><b>OTHER PERMITS, INSPECTIONS, and REQUIREMENTS</b></u></center>
<br />
A copy of the approved plans (if applicable) must be kept on the job site at all times and available for the Inspector. This permit must be placed in a door or a window so as to be clearly visible from the street. Plumbing and electrical permits are also required if the project involves these trades. Please contact the following agencies to obtain applications, schedule inspections, or if you have any questions as to the necessity of these permits.
<ul>
<li>For plumbing permits contact Hamilton County Public Health, 250 William Howard Taft, Cincinnati, OH 45219, phone: (513) 946-7800.</li>
<li>For electric permits contact the Inspection Bureau Incorporated, 250 W. Court St., Ste. 125 W, Cincinnati, OH 45202, phone: (513) 381-6080. A Bonding Inspection is required by IBI prior to placing the concrete in any footing of a new residential building, addition, swimming pool or other structure if the structure is electrified.</li>
</ul>

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
