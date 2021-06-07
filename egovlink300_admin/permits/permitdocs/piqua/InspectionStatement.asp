<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<link rel="stylesheet" type="text/css" href="style.css">
	</head>
	<%
		strTitle = "INSPECTION STATEMENT"
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
		The changes or work authorized by the permit issued for this project have been inspected as noted. The issuance of this Inspection Statement acknowledges the completeness and appropriateness of the changes and or work inspected to date.
		<hr />
		<!--#include file="header.asp"-->
		<table>
			<tr>
				<th colspan="4">INSPECTIONS</th>
			</tr>
			<tr>
				<td width="50%">INSPECTION</td>
				<td>STATUS</td>
				<td>DATE</td>
				<td>INSPECTOR</td>
			</tr>
			<%
				Do While Not oRsI.EOF
				%>
					<tr>
						<td><%=oRsI("inspection")%></td>
						<td><%=oRsI("inspectionstatus")%></td>
						<td><%=oRsI("inspecteddate")%></td>
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
			<tr>
				<td colspan="4">
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
