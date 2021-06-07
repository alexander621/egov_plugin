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
	<body onLoad="window.print();">
Date:<%=date()%><br />
<br />
<table>
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
			Re:<%=oRs("permittypedesc")%><br />
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
	</tr>
</table>
<br />
<br />
<br />
<br />
Dear Permit Applicant:<br />
<br />
Upon reviewing the recently submitted permit application for the above project, it was found that additional information is needed to enable us to determine if the proposed improvements meet the life safety provisions of the Residential Code of Ohio and/or the various aspects of the Wyoming Zoning Code. Please provide the following information or modify the plans to address the following matters.<br />
<table>
	<tr>
		<td width="50%">
			Plan Review Comments:
			<%=sReviewNotes%>
		</td>
		<td width="50%">
			Code Reference:<br />
		</td>
	</tr>
</table>
<br />
Three complete sets of plans and specifications should be submitted. In accordance with Section 110 of the Residential Code of Ohio and/or Chapter 1135 of the Wyoming Codified Ordinances, you may appeal this decision within thirty days. If you have any questions, please feel free to contact the City of Wyoming Community Development Department at (513) 821-7600.<br />
Respectfully,<br /><br />
Megan Statt Blake, Community Development Director<br />
Daniel Bly, Residential Building Official and Plans Examiner
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
