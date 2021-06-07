<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<link rel="stylesheet" type="text/css" href="style.css">
		<style>
			body, table
			{
				font-size:10px;
			}
		</style>
	</head>
	<%
		strTitle = "PERMIT"
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
		<br />
		<table>
			<tr>
				<th colspan="2">INSPECTIONS</th>
			</tr>
			<tr>
				<td width="50%">INSPECTION</td>
				<td>INSPECTOR</td>
			</tr>
			<%
				Do While Not oRsI.EOF
				%>
					<tr>
						<td><%=oRsI("inspection")%></td>
						<td>
							<%=oRsI("inspector")%>
							<% if trim(oRsI("inspectorphone")) <> "" then response.write " - " & FormatPhoneNumber(oRsI("inspectorphone"))%>
						</td>
					</tr>
					<%oRsI.MoveNext
				loop
				oRsI.Close
				Set oRsI = Nothing
			%>
		</table>
		A code compliance review of the application and construction documents submitted has resulted in the release of this Permit for the above referenced project. The changes or work authorized by this permit are as noted above and in the referenced documents. The inspections necessary to the proposed changes or work to be completed are also identified above. It is the responsibility of the owner or applicant or their agent to contact the appropriate office to schedule the necessary inspections. All inspections conducted and the final approval of the changes or work authorized will be subject to all items being completed in accordance with the Plan Review notes and the permit application and construction document submittals provided in support of this permit request. Please contact Miami County Building Regulations (937) 440-8075 concerning building, electrical, or mechanical permit requirements that may be applicable to the work being performed. For information regarding plumbing permit requirements that may be applicable to the work being performed contact the City of Piqua Health Department (937) 778-2060.
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
		<%=sReviewNotes%>


		<% 
			oRs.Close
			Set oRs = Nothing
			oRsA.Close
			Set oRsA = Nothing
		%>
	</body>
</html>
