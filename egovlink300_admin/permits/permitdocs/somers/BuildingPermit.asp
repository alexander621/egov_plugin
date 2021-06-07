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
			h2 {font-weight:normal;}
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

			.noleftborder { border-left:none; }
			.notopborder { border-top:none; }
		</style>
	</head>
	<%
		intPermitID = request.querystring("permitid")
	%>
	<!--#include file="../getdata.asp"-->
	<body onLoad="window.print();">
		<% strTITLE="BUILDING PERMIT" %>
		<!--#include file="permittop.asp"-->
			<tr>
				<td><b>PROJECT INFORMATION</b></td>
				<td>Valuation: <%=FormatCurrency(oRs("JobValue"),2)%></td>
			</tr>
			<tr>
				<td>Use Type: <%=oRs("usetype")%></td>
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
		<table>
			<tr>
				<td width="80%" class="notopborder"><b>FEES</b></td>
				<td class="notopborder">Total Fees: <%=FormatCurrency(oRs("FeeTotal"),2)%></td>
			</tr>
		</table>
		<table>
			<tr>
				<td class="notopborder">Payment Date:</td>
				<td class="notopborder noleftborder"><%=sPaymentDate1%></td>
				<td class="notopborder">Method:</td>
				<td class="notopborder noleftborder"><%=sMethod1%></td>
				<td class="notopborder">Amount:</td>
				<td class="notopborder noleftborder"><%=sAmount1%></td>
			</tr>
			<tr>
				<td>Payment Date:</td>
				<td class="noleftborder"><%=sPaymentDate2%></td>
				<td>Method:</td>
				<td class="noleftborder"><%=sMethod2%></td>
				<td>Amount:</td>
				<td class="noleftborder"><%=sAmount2%></td>
			</tr>
		</table>
		<table>
			<tr>
				<td width="80%" align="center" class="notopborder"><b>*If there are more than two payments please see the Invoice Summary*</b></td>
				<td class="notopborder">Total Payments: <%=sTotalPaid%></td>
			</tr>
		</table>
		<b>ADDITIONAL COMMENTS:</b><br />
		<%=oRs("PermitNotes")%>
		<br />
		<br />
		<br />
		<center><b>
		ALL WORK SHALL BE IN COMPLIANCE WITH ALL STATE AND LOCAL REGULATIONS AS APPROVED BY THE
		<br />
		<br />
		SOMERS BUILDING DEPARTMENT
		</b></center>
		<br />
		<br />
		The approval of plans or drawings and specifications or data in accordance with this rule is invalid if construction, erection, alteration, or other work upon the building has not commenced within twelve months of the approval of the plans or drawings and specifications. Two extensions shall be granted for six months each if requested by the owner at least ten days in advance of the expiration of the approval.
		<br />
		<br />
		<br />
		<br />
		Approving Official ___________________________________________________ Date Issued ___________________________
		




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
