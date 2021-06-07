<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitrevenuereportexport.asp
' AUTHOR: Steve Loar
' CREATED: 01/28/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permit payments by payment media (cash, check, cc)
'
' MODIFICATION HISTORY
' 1.0   01/28/2009	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromPaymentDate, sToPaymentDate, sStreetNumber, sStreetName, sPermitNo
Dim sPayor, sInvoiceNo, sDisplayDateRange, sApplicant, iWaivedPick
Dim sSql, oRs, iRowCount, dTotalAmount, dSubTotalAmount, iOldReportingFeeTypeId, sFeeType
Dim sDate, sFrom, sWhere, sSum

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Permit_Revenue_Report_" & sDate & ".xls"


If request("waivedpick") <> "" Then
	iWaivedPick = clng(request("waivedpick"))
	sSearch = sSearch & " AND I.allfeeswaived = " & request("waivedpick")
Else
	iWaivedPick = 0
	sSearch = sSearch & " AND I.allfeeswaived = 0 "
End If 

' Handle payment date range. always want some dates to limit the search
If request("topaymentdate") <> "" And request("frompaymentdate") <> "" Then
	sFromPaymentDate = request("frompaymentdate")
	sToPaymentDate = request("topaymentdate")
	If iWaivedPick = 0 Then 
		sSearch = sSearch & " AND (J.paymentdate >= '" & request("frompaymentdate") & "' AND J.paymentdate < '" & DateAdd("d",1,request("topaymentdate")) & "' ) "
	Else
		sSearch = sSearch & " AND (I.invoicedate >= '" & request("frompaymentdate") & "' AND I.invoicedate < '" & DateAdd("d",1,request("topaymentdate")) & "' ) "
	End If 
	sDisplayDateRange = "From: " & request("frompaymentdate") & " &nbsp;To: " & request("topaymentdate")
End If 

' handle address pick
If request("residentstreetnumber") <> "" Then 
'	sStreetNumber = request("residentstreetnumber")
	sSearch = sSearch & "AND A.residentstreetnumber = '" & dbsafe(request("residentstreetnumber")) & "' "
End If 
If request("streetname") <> "" And request("streetname") <> "0000" Then 
	sStreetName = request("streetname")
	sSearch = sSearch & " AND (A.residentstreetname = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix = '" & dbsafe(sStreetName) & "' "
	sSearch = sSearch & " OR A.residentstreetprefix + ' ' + A.residentstreetname + ' ' + A.streetsuffix + ' ' + A.streetdirection = '" & dbsafe(sStreetName) & "' )"
End If 

' handle the permit number
If request("permitno") <> "" Then 
	sPermitNo = Trim(request("permitno"))
	sSearch = sSearch & BuildPermitNoSearch( sPermitNo )	' in permitcommonfunctions.asp
End If 

If request("invoiceno") <> "" Then 
	sInvoiceNo = CLng(request("invoiceno"))
	sSearch = sSearch & " AND I.invoiceid = " & sInvoiceNo
End If 

If request("applicant") <> "" Then 
	sApplicant = request("applicant")
	sSearch = sSearch & " AND ( C.company LIKE '%" & dbsafe(sApplicant) & "%' OR C.firstname LIKE '%" & dbsafe(sApplicant) & "%' OR C.lastname LIKE '%" & dbsafe(sApplicant) & "%' ) "
End If 

If request("permitcategoryid") <> "" Then
	If CLng(request("permitcategoryid")) > CLng(0) Then
		sSearch = sSearch & " AND P.permitcategoryid = " & request("permitcategoryid")
	End If 
End If 

If request("permitlocation") <> "" Then
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If

dTotalAmount = CDbl(0.00)
dSubTotalAmount = CDbl(0.00)
iOldReportingFeeTypeId = CLng(-1)
iRowCount = 0

If clng(iWaivedPick) = clng(0) Then
	sDate = "J.paymentdate"
	sFrom = "egov_class_payment J, egov_accounts_ledger AL,"
	sWhere = " AND II.permitfeeid = AL.permitfeeid AND II.invoiceid = AL.invoiceid AND J.paymentid = AL.paymentid"
	sSum = "AL.amount"
Else
	sDate = "I.invoicedate"
	sFrom = ""
	sWhere = ""
	sSum = "I.totalamount"
End If 

sSql = "SELECT ISNULL(II.feereportingtypeid,0) AS feereportingtypeid, I.permitid, I.invoiceid, " & sDate & " AS paymentdate, "
sSql = sSql & " ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, "
sSql = sSql & " SUM(" & sSum & ") AS amount "
sSql = sSql & " FROM egov_permitinvoiceitems II, egov_permitinvoices I, "
sSql = sSql & sFrom & " egov_permitaddress A, egov_permits P, egov_permitcontacts C, egov_permitlocationrequirements R "
sSql = sSql & " WHERE II.invoiceid = I.invoiceid AND I.orgid = " & session("orgid") & " AND I.isvoided = 0 "
sSql = sSql & " AND I.permitid = C.permitid AND C.isapplicant = 1 AND P.permitlocationrequirementid = R.permitlocationrequirementid "
sSql = sSql & sSearch
sSql = sSql & " AND I.permitid = P.permitid AND P.isvoided = 0 AND A.permitid = I.permitid " & sWhere
sSql = sSql & " GROUP BY II.feereportingtypeid, I.permitid, I.invoiceid, " & sDate & ", P.permitlocation, R.locationtype, dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) "
sSql = sSql & " ORDER BY II.feereportingtypeid, I.permitid, I.invoiceid, " & sDate 


Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	response.write vbcrlf & "<html>"

	response.write vbcrlf & "<style>  "
	response.write " .moneystyle "
	response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
	response.write vbcrlf & "</style>"

	response.write vbcrlf & "<body>"
'	response.write sSql & "<br /><br />"

	response.write vbcrlf & "<table cellpadding=""4"" cellspacing=""0"" border=""1"">"
	response.write vbcrlf & "<tr><th>Fee Type</th><th>Permit #</th><th>Invoice #</th><th>Payment<br />Date</th><th>Address/Location</th><th>Amount</th></tr>"
	response.flush

	Do While Not oRs.EOF
		If iOldReportingFeeTypeId <> CLng(oRs("feereportingtypeid")) Then
			If iOldReportingFeeTypeId <> CLng(-1) Then
				response.write vbcrlf & "<tr><td colspan=""4"">&nbsp;</td><td>" & sFeeType & " Total</td>"
				response.write "<td align=""right"" class=""moneystyle"">" & dSubTotalAmount & "</td></tr>"
				dSubTotalAmount = CDbl(0.00)
			End If 
			iRowCount = 0
			iOldReportingFeeTypeId = CLng(oRs("feereportingtypeid"))
			' Print permit fee type name
			sFeeType = GetFeeReportingType( iOldReportingFeeTypeId )
			response.write "<tr><td colspan=""6"">&nbsp;" & sFeeType & "</td></tr>"
			response.flush
		End If 

		dTotalAmount = dTotalAmount + CDbl(oRs("amount"))
		dSubTotalAmount = dSubTotalAmount + CDbl(oRs("amount"))
		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr>"
		response.write "<td>&nbsp;</td>"
		response.write "<td align=""center"">" & GetPermitNumber( oRs("permitid") ) & "</td>"
		response.write "<td align=""center"">" & oRs("invoiceid") & "</td>"
		response.write "<td align=""center"" nowrap=""nowrap"">" & FormatDateTime(oRs("paymentdate"),2) & "</td>"

		'response.write "<td nowrap=""nowrap"">&nbsp;" & oRs("permitaddress") & "</td>"
		response.write "<td>"
			Select Case oRs("locationtype")
				Case "address"
					response.write oRs("permitaddress")

				Case "location"
					response.write oRs("permitlocation")

				Case Else
					response.write ""

			End Select  
			response.write "</td>"

		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(CDbl(oRs("amount")),2,,,0) & "</td>"
		response.flush
		oRs.MoveNext
	Loop 
	' last sub total Row
	If iOldReportingFeeTypeId <> CLng(-1) Then
		response.write vbcrlf & "<tr><td colspan=""4"">&nbsp;</td><td>" & sFeeType & " Total</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(CDbl(dSubTotalAmount),2,,,0) & "</td></tr>"
	End If 
	response.flush
	' Grand Total Row
	response.write vbcrlf & "<tr><td colspan=""4"">&nbsp;</td><td>Grand Total</td>"
	response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(CDbl(dTotalAmount),2,,,0) & "</td></tr>"
	response.write "</table>"
	response.flush

	response.write vbcrlf & "</table></body></html>"
'Else
'	response.write sSql & "<br /><br />"
End If 

oRs.Close
Set oRs = Nothing 

%>


<!-- #include file="permitcommonfunctions.asp" //-->
<!-- #include file="../includes/common.asp" //-->
