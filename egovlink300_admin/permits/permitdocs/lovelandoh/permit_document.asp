<!--#include file="../../../includes/common.asp" -->
<!--#include file="../../permitcommonfunctions.asp" -->
<html>
	<head>
		<style>
			*
			{
				font-size:12px;
			}
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
			table.border td, table.border th
			{
				border-top: 1px solid black;
				border-left: 1px solid black;
				padding: 5px;
			}
			table.border td:last-child, table.border th:last-child
			{
				border-right: 1px solid black;
			}
			table.border tr:last-child td
			{
				border-bottom: 1px solid black;
			}
			/*table.border tr:first-child td
			{
				border: none;
			}*/
			table.noborder td, table.noborder th
			{
				border:none;
			}
			table.notopborder tr:first-child td
			{
				border-top:0;
			}

		</style>
	</head>
	<%
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<body onload="window.print();">
		<table>
			<tr>
				<td align="right" valign="top"><img src="loveland-logo-2.png" height="100" /></td>
				<td align="center" valign="top" width="50%">
					<h2 style="font-size:16px;">BUILDING/ZONING PERMIT</h2>
				</td>
				<td valign="top">
				<b>City of Loveland<br />
				Building & Zoning</b><br />
				120 W. Loveland Ave.<br />
				Loveland, Ohio 45140<br />
				Office: 513.707.1447<br />
				Fax: 513.583.3032<br />
				www.lovelandoh.com
				</td>
			</tr>
		</table>
		<center>
		<b>Permit No: <%=strPermitNumber%>
		<br />
		<u>**APPROVED PLANS ARE REQUIRED TO BE ON SITE FOR EVERY INSPECTION**
		<br />
		FOR INSPECTIONS CALL (513) 707-1447 10:00am TO 2:00pm MONDAY-FRIDAY</u></b>
		</center>
		<br />
		<table style="border: 1px solid black;">
			<tr>
				<td width="50%" valign="top">
					<table>
						<td valign="top"><b>Job Site Address:</b></td>
						<td>
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
					</table>
				</td>
				<td width="50%" valign="top">
					<table>
						<td valign="top"><b>Property Owner:</b></td>
						<td>
				<%
					response.write oRs("ListedOwner")
					%>
						</td>
					</table>
				</td>
			</tr>
			<tr>
				<td width="50%" valign="top">
					<table>
						<td valign="top"><b>Primary Contractor:</b></td>
						<td>
						<%
							if oRs("isprimarycontractor") then
								If oRs("confirstname") <> "" Then 
									response.write oRs("confirstname") & " " & oRs("conlastname") & "<br />"
								End If 
								If oRs("concompany") <> "" Then 
									response.write oRs("concompany") & "<br />" 
								End If 
								If Trim(oRs("conaddress")) <> "" Then 
									response.write oRs("conaddress") & "<br />" 
								End If 
								If Trim(oRs("concity")) <> "" Then 
									response.write oRs("concity") & ", " & oRs("constate") & " " & oRs("conzip") & "<br />"
								End If 
								If Not IsNull(oRs("conphone")) And Trim(oRs("conphone")) <> "" Then 
									response.write FormatPhoneNumber( oRs("conphone") ) 
								End If 
							else
								If oRs("appfirstname") <> "" Then 
									response.write oRs("appfirstname") & " " & oRs("applastname") & "<br />"
								End If 
								If oRs("appcompany") <> "" Then 
									response.write oRs("appcompany") & "<br />" 
								End If 
								If Trim(oRs("appaddress")) <> "" Then 
									response.write oRs("appaddress") & "<br />" 
								End If 
								If Trim(oRs("appcity")) <> "" Then 
									response.write oRs("appcity") & ", " & oRs("appstate") & " " & oRs("appzip") & "<br />"
								End If 
								If Not IsNull(oRs("appphone")) And Trim(oRs("appphone")) <> "" Then 
									response.write FormatPhoneNumber( oRs("appphone") ) 
								End If 
							end if
						%>
						</td>
					</table>
				</td>
				<td width="50%" valign="top">
					<table>
						<td valign="top"><b>Plans By:</b></td>
						<td>
						<%
						response.write sPlansBy
						%>
						</td>
					</table>
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<b>Primary Contact:</b><br />
					&nbsp;
				</td>
			</tr>
		</table>
		<table class="border notopborder">
			<tr>
				<td><b>Construction Type: <%=oRs("constructiontype")%></b></td>
				<td><b>Use Group:</b> <%=oRs("usegroupcode")%></td>
				<td><b>Occupants:</b> <%=oRs("occupants")%></td>
			</tr>
			<tr>
				<td colspan="3"><b>Approved As:</b> <%=oRs("approvedas")%></td>
			</tr>
		</table>
		<br />

		<b>LOCAL TAXES:</b><br />
		<br />
		The City of Loveland has a 1% local income tax which is levied on individuals, business’ net profits and employee wages. Businesses or individuals doing sales, work, or services within the city limits are required to file an annual tax return and remit the appropriate tax due. All forms and instructions regarding these tax returns can be found at www.ritaohio.com. Please contact 513-707-1452 with questions.
		<br />
		<br />
		<b>ALL WORK SHALL BE IN COMPLIANCE WITH APPLICABLE BUILDING & ZONING REGULATIONS AS APPROVED BY THE BUILDING & ZONING DEPARTMENT</b><br />
		The approval of plans or drawings and specifications or data in accordance with this rule is invalid if construction, erection, alteration, or other work upon the building has not commenced within twelve months of the approval of the plans or drawings and specifications. One extension shall be granted for an additional twelve-month period if requested by the owner at least ten days in advance of the expiration of the approval and upon payment of a fee not to exceed one hundred dollars. If in the course of construction, work is delayed or suspended for more than six months, the approval of plans or drawings and specifications or data is invalid. Two extensions shall be granted for six months each if requested by the owner at least ten days in advance of the expiration of the approval and upon payment of a fee for each extension of not more than one hundred dollars.
		<br />
		<br />
		Approving Official___________________________________ Date Issued____<%if not isnull(oRs("issueddate")) then response.write FormatDateTime(oRs("issueddate"),2)%>_____
		<br />
		<br />
		<table class="border">
			<tr>
				<td><b>Fees</b></td>
				<td>Permit: <%=dPermitFee%></td>
				<td>State Fee: <%=dBBSFee%></td>
				<td>Zoning Fee: <%=dZoneFee%></td>
			</tr>
			<tr>
				<td><b>Other Fees</b></td>
				<td>Road Impact: <%=dRoadImpactFee%></td>
				<td>Water Impact: <%=dWaterImpactFee%></td>
				<td>Recreation Impact: <%=dRecImpactFee%></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td>Water Meter: <%=dWaterMeterFee%></td>
				<td>&nbsp;</td>
				<td>Total Fees: <%=sTotalFees%></td>
			</tr>
			<tr>
				<td>Payment Date: <%=sPaymentDate1%></td>
				<td>Method: <%=sMethod1%></td>
				<td colspan="2">Amount: <%=sAmount1%></td>
			</tr>
			<tr>
				<td>Payment Date: <%=sPaymentDate2%></td>
				<td>Method: <%=sMethod2%></td>
				<td colspan="2">Amount: <%=sAmount2%></td>
			</tr>
			<tr>
				<td colspan="3"><b>*If there are more than two payments please see the Invoice Summary*</b></td>
				<td>Total Payments: <%=sTotalPaid%></td>
			</tr>
		</table>
		<br />
		<table style="border: 1px solid black;">
			<tr>
				<td width="5%" align="right">County:</td>
				<td width="18%" style="border-bottom: 1px solid black;"><%=oRs("county")%></td>
				<td width="9%" align="right">Parcel ID:</td>
				<td width="18%" style="border-bottom: 1px solid black;"><%=oRs("parcelidnumber")%></td>
				<td width="9%" align="right">Floor Area:</td>
				<td width="18%" style="border-bottom: 1px solid black;"><%=sDimensions%></td>
				<td width="5%" align="right">Valuation:</td>
				<td width="18%" style="border-bottom: 1px solid black;"><%=FormatCurrency(oRs("jobvalue"),2)%></td>
			</tr>
			<tr>
				<td colspan="8">
					<table>
						<td width="18%">Description of Work</td>
						<td style="border-bottom: 1px solid black;"><%=oRs("descriptionofwork")%></td>
					</table>
				</td>
			</tr>
		</table>
		<br />
		For Electric Permits Contact Inspection Bureau Inc. at 513.381.6080.
		<br />
		For Plumbing Permits Contact Hamilton County Plumbing Dept. at 513.946.7800.
		<br />
		For Sewer Permits Contact Metropolitan Sewer District at 513.352.4900.
		<br />
		Additional Permits Required:<br />
		___ HVAC<br />
		___ Electric<br />
		___ Plumbing



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
