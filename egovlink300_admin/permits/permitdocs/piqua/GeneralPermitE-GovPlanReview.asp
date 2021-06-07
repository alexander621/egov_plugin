<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<link rel="stylesheet" type="text/css" href="style.css">
		<style>
			body
			{
				font-size:12px;
			}
		</style>
	</head>
	<%
		strTitle = "PLAN REVIEW STATEMENT"
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<body onload="window.print();">
		<!--#include file="header.asp"-->
		<table>
			<tr>
				<th width="50%">APPLICANT INFORMATION</th>
				<th width="50%">PRIMARY CONTACT</th>
			</tr>
			<tr>
				<td>
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
				<td>
					<table class="noborder">
						<%
							'if not isnull(oRs("confirstname")) then
								'strConName = oRs("confirstname") & " " & oRs("conlastname")
								'strConPhone = FormatPhoneNumber(oRs("conphone"))
							'else
								strConName = oRs("appfirstname") & " " & oRs("applastname")
								strConPhone = FormatPhoneNumber(oRs("appphone"))
							'end if
							strConCell = FormatPhoneNumber(oRs("appcell"))
							strConFax = FormatPhoneNumber(oRs("appfax"))
							strConEmail = oRs("appemail")

						%>
						<tr>
							<td>Name: <%=strConName%></td>
						</tr>
						<tr>
							<td>Phone: <%=strConPhone%></td>
						</tr>
						<tr>
							<td>Cell: <%=strConCell%></td>
						</tr>
						<tr>
							<td>Fax: <%=strConFax%></td>
						</tr>
						<tr>
							<td>Email: <%=strConEmail%></td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		<br />
		<table>
			<tr>
				<th colspan="2">PROJECT INFORMATION</th>
			</tr>
			<tr>
				<td>Use Type: <%=oRs("UseTYpe")%></td>
				<td>Use Class: <%=oRs("useclass")%></td>
			</tr>
			<tr>
				<td colspan="2">Description of Work: <%=oRs("descriptionofwork")%></td>
			</tr>
			<tr>
				<td>Work Scope: <%=oRs("workscope")%></td>
				<td>Work Class: <%=oRs("workclass")%></td>
			</tr>
		</table>
		<br />
		<table>
			<tr>
				<th colspan="3">REFERENCE DOCUMENTS</th>
			</tr>
			<tr>
				<td>DATE SUBMITTED</td>
				<td>FILE NAME</td>
				<td>DESCRIPTION</td>
			</tr>
			<% do while not oRsA.EOF %>
				<tr>
					<td><%=oRsA("dateadded")%></td>
					<td><%=oRsA("attachmentname")%></td>
					<td><%=oRsA("description")%></td>
				</tr>
			<%	oRsA.MoveNext
			loop%>
		</table>
		A code compliance review of the proposed changes or work referenced above has been completed. The review of the application and construction document submittals has resulted in the issuance of this Plan Review statement. The Plan Review notes provided indicate the appropriateness of the application and construction documents with regards to the conformance of the proposed changes or work with the adopted community zoning, stormwater, water distribution, electric distribution, and sanitary sewer standards. Please review the Plan Review notes provided and respond accordingly by revising the construction documents and or by compiling the material necessary to address the concern noted, and submit the revised drawings and or additional information to the Development Office.
		<br />
		<br />
		Please contact the plan review person listed next to the pertinent review topic to inquiry with questions or concerns related to the Plan Review notes specific to that particular review topic. Direct all other inquiries related to the Permit request to the Development Office.
		<br />
		<br />
		Note - The permit request is NOT APPROVED AND WORK IS NOT AUTHORIZED TO COMMENCE until all of the Plan Review notes provided herein have been satisfactorily addressed. A Permit with a permit application status of “Approved” will be issued to the applicant upon the proposed changes or work shown on the application and or construction document submittals being found to be appropriate to and in conformance with the adopted community standards.
		<hr />
		<!--#include file="header.asp"-->
		<table>
			<tr>
				<th colspan="4">PLAN REVIEW NOTES</th>
			</tr>
			<!--tr>
				<td>REVIEW</td>
				<td>STATUS</td>
				<td>DATE</td>
				<td>REVIEWER</td>
			</tr-->
			<% Do While Not oRsPR.EOF %>
				<!--tr>
					<td><%=oRsPR("permitreviewtype")%></td>
					<td><%=oRsPR("reviewstatus")%></td>
					<td><%=oRsPR("reviewed")%></td>
					<td>
						<%=oRsPR("reviewer")%>
						<% if trim(oRsPR("reviewerphone")) <> "" then response.write " - " & FormatPhoneNumber(oRsPR("reviewerphone"))%>
					</td>
				</tr-->
			<%	oRsPR.MoveNext
			loop%>
			<tr><td colspan="4" style="border-bottom:0;"><%=sReviewNotes%></td></tr>
			<tr>
				<td colspan="4" style="border-top:0;">
					<br />
					<br />
					<br />
					<br />
					<br />
					<div style="border: 2px solid black;padding:8px;">
					Signature of Authorized Official: 
					</div>
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
