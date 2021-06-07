<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: sewerfeereportexport.asp
' AUTHOR: Steve Loar
' CREATED: 09/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of sewer connection fees, dumped to excel
'
' MODIFICATION HISTORY
' 1.0   09/30/2008	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Sewer_Connection_Fee_Report_" & sDate & ".xls"

	Dim sSearch, sRptTitle

	sSearch = session("sSql")
	
	sRptTitle = vbcrlf & "<tr><th>Sewer Connection Fee Report</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

	DisplaySewerConnectionFees sSearch, sRptTitle

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Sub DisplaySewerConnectionFees( sSearch, sRptTitle )
'--------------------------------------------------------------------------------------------------
Sub DisplaySewerConnectionFees( ByVal sSearch, sRptTitle )
	Dim sSql, oRs, iRowCount, dSewerFeeTotal, dFeesTotal

	iRowCount = 0
	dSewerFeeTotal = CDbl(0.00)
	dFeesTotal = CDbl(0.00)

	sSql = "SELECT P.permitid, P.applieddate, P.issueddate, I.invoiceid, ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, ISNULL(C.company,'') AS company, C.contacttype, "
	sSql = sSql & " ISNULL(P.feetotal,0.00) AS feetotal, F.permitfeeprefix, F.permitfee, ISNULL(F.feeamount,0.00) AS feeamount "
	sSql = sSql & " FROM egov_permits P, egov_permitcontacts C, egov_permitfees F, egov_permitinvoiceitems I, egov_permitinvoices II, egov_permitfeereportingtypes R "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.permitid = C.permitid AND C.isapplicant = 1 AND P.permitid = F.permitid "
	sSql = sSql & " AND F.feereportingtypeid = R.feereportingtypeid AND R.issewerconnection = 1 AND I.permitid = P.permitid AND F.permitfeeid = I.permitfeeid "
	sSql = sSql & " AND I.invoiceid = II.invoiceid AND II.permitid = P.permitid AND II.isvoided = 0 " & sSearch
	sSql = sSql & " ORDER BY P.applieddate, I.invoiceid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<html><body><table border=""1"" cellpadding=""2"">"
		'response.write sRptTitle
		response.write vbcrlf & "<tr height=""30""><th>Open Date</th><th>Close Date</th><th>Permit #</th><th>Fee Cat</th><th>Description</th><th>Fee</th><th>Total Amount</th><th>Invoice</th><th>Applicant</th></tr>"
		response.flush

		Do While Not oRs.EOF
			iRowCount = iRowCount + 1
			response.write vbcrlf & "<tr height=""26"">"

			' Open Date
			response.write "<td align=""center"">" & FormatDateTime(oRs("applieddate"),2) & "</td>"

			' Close date
			response.write "<td align=""center"">" 
			If IsNull(oRs("issueddate")) Then 
				response.write "&nbsp;"
			Else
				response.write FormatDateTime(oRs("issueddate"),2) 
			End If 
			response.write "</td>"

			'Permit Number
			response.write "<td align=""center"">&nbsp;" & GetPermitNumber( oRs("permitid") ) & "</td>"

			' Fee Category
			response.write "<td align=""center"">&nbsp;" & oRs("permitfeeprefix") & "</td>"

			' Fee Desctiption 
			response.write "<td align=""left"" width=""600"">" & oRs("permitfee") & "</td>"

			' Fee Amount
			response.write "<td align=""right"">&nbsp;" & FormatNumber(oRs("feeamount"),2) & "</td>"
			dSewerFeeTotal = dSewerFeeTotal + CDbl(oRs("feeamount"))

			' Total of fees for the permit
			response.write "<td align=""right"">&nbsp;" & FormatNumber(oRs("feetotal"),2) & "</td>"
			dFeesTotal = dFeesTotal + CDbl(oRs("feetotal"))

			' Invoice Number
			response.write "<td align=""center"">" & oRs("invoiceid") & "</td>"

			' Applicant
			response.write "<td align=""left"" width=""400"">"
			If oRs("firstname") <> "" Then 
				response.write oRs("firstname") & " " & oRs("lastname")
			Else
				response.write oRs("company")
			End If 
			response.write "</td>"

			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop
		' Totals row
		response.write vbcrlf & "<tr class=""totalrow"" height=""26""><td colspan=""5"">&nbsp;</td><td align=""right""><b>&nbsp;" & FormatNumber(dSewerFeeTotal,2) & "</b></td><td align=""right""><b>&nbsp;" & FormatNumber(dFeesTotal,2) & "</b></td><td colspan=""2"">&nbsp;</td></tr>"
		response.write vbcrlf & "</table></body></html>"
		response.flush
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


%>

<!-- #include file="permitcommonfunctions.asp" //-->


