<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitmonthlyreportexport.asp
' AUTHOR: Steve Loar
' CREATED: 11/13/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permits issued, dumped to excel
'
' MODIFICATION HISTORY
' 1.0   11/13/2008	Steve Loar - INITIAL VERSION
' 1.1	05/20/2010	Steve Loar - Added pick of incremental job value and total job value for cost estimate
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStartMonth, iStartYear, sStartDate, sYearStart, iEndMonth, iEndYear, sEndDate, sYearEnd
Dim sSql, oRs, iRowCount, sReportGroup, dCostEstimateSubTotal, dCostEstimateTotal, sClass
Dim iSubTotalUnits, iTotalUnits, dZoneSubTotal, dZoneTotal, dBBSSubTotal, dBBSTotal
Dim dPenSubTotal, dPenTotal, dCertSubTotal, dCertTotal, dPermitSubTotal, dPermitTotal
Dim iYTDUnits, iPreviousUnits, dYTDCostEstimate, dPreviousCostEstimate, dYTDPermitFees
Dim dPreviousPermitFees, dYTDZone, dPreviousZone, dYTDBBS, dPreviousBBS, dYTDPen
Dim dPreviousPen, dYTDCert, dPreviousCert, iInclude, bUsePermitJobValue, iIncludeJobValue

sReportGroup = "None"
iRowCount = 0
dCostEstimateSubTotal = CDbl(0.00)
dCostEstimateTotal = CDbl(0.00)
iSubTotalUnits = CLng(0)
iTotalUnits = CLng(0)
dZoneSubTotal = CDbl(0.00)
dZoneTotal = CDbl(0.00)
dBBSSubTotal = CDbl(0.00)
dBBSTotal = CDbl(0.00)
dPenSubTotal = CDbl(0.00)
dPenTotal = CDbl(0.00)
dCertSubTotal = CDbl(0.00)
dCertTotal = CDbl(0.00)
dPermitSubTotal = CDbl(0.00)
dPermitTotal = CDbl(0.00)
sClass = ""


' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Permit_Monthly_Report_" & sDate & ".xls"

If request("selmonth") <> "" Then
	iStartMonth = clng(request("selmonth"))
Else
	iStartMonth = clng(Month(Date))   ' this month
End If 
If request("selyear") <> "" Then
	iStartYear = clng(request("selyear"))
Else
	iStartYear = clng(Year(Date)) ' This Year
End If 
sStartDate = iStartMonth & "/01/" & iStartYear
sYearStart = "01/01/" & iStartYear
'sYearEnd = "01/01/" & (iStartYear + 1)

If iStartMonth < clng(12) Then
	iEndMonth = iStartMonth + 1
	iEndYear = iStartYear
Else
	iEndMonth = 1
	iEndYear = iStartYear + 1
End If 
sEndDate = iEndMonth & "/01/" & iEndYear
sYearEnd = sEndDate

If request("include") <> "" Then
	iInclude = request("include")
Else
	iInclude = 0
End If 

If clng(iInclude ) < clng(2) Then
	sIsVoided = " AND P.isvoided = " & iInclude
Else
	sIsVoided = ""
End If 

If request("includejobvalue") <> "" Then
	iIncludeJobValue = request("includejobvalue")
	If clng(iIncludeJobValue) = clng(0) Then 
		bUsePermitJobValue = False 
	Else
		bUsePermitJobValue = True 
	End If 
Else
	bUsePermitJobValue = False 
End If



sSql = "SELECT P.permitid, P.issueddate, P.jobvalue, ISNULL(P.residentialunits,0) AS residentialunits, "
sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 0 AS isold, U.reportgroup, "
sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitusetypes U "
sSql = sSql & " WHERE P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "' "
sSql = sSql & " AND A.permitid = P.permitid AND P.usetypeid = U.usetypeid AND P.orgid = " & session("orgid")
sSql = sSql & sIsVoided
sSql = sSql & " UNION ALL "
sSql = sSql & " SELECT P.permitid, P.issueddate, P.jobvalue, 0 AS residentialunits, "
sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 1 AS isold, U.reportgroup, "
sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitinvoices I, egov_permitusetypes U "
sSql = sSql & " WHERE A.permitid = P.permitid AND I.permitid = P.permitid AND I.invoicedate > P.issueddate AND P.issueddate < '" & sStartDate & "' "
sSql = sSql & " AND I.invoicedate >= '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' AND P.usetypeid = U.usetypeid AND P.orgid = " & session("orgid")
sSql = sSql & " AND I.paymentid IS NOT NULL AND I.isvoided = 0 AND I.allfeeswaived = 0 " & sIsVoided
sSql = sSql & " ORDER BY U.reportgroup, P.issueddate, P.permitid, isold"
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	response.write vbcrlf & "<html><body>"
	response.write vbcrlf & "<table cellpadding=""4"" cellspacing=""0"" border=""1"">"
	response.write vbcrlf & "<tr height=""30""><th>Permit #</th><th>Issued Date</th><th>Description<br />of Work</th><th>Address</th><th>Residential Units</th><th>Est. Cost</th>"
	response.write "<th>Permit</th><th>Zone Fee</th><th>BBS Fee</th></tr>"   '<th>Rev./Pen</th><th>C of O</th></tr>"
	response.flush
	Do While Not oRs.EOF
		If sReportGroup <> oRs("reportgroup") Then
			If sReportGroup <> "None" Then
				' Print out a subTotalLine
				response.write vbcrlf & "<tr height=""20""><td colspan=""4"">&nbsp;</td>"
				response.write "<td align=""center""><b>" & iSubTotalUnits & "</b></td>"
				response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dCostEstimateSubTotal,2) & "</b></td>"
				response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dPermitSubTotal,2) & "</b></td>"
				response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dZoneSubTotal,2) & "</b></td>"
				response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dBBSSubTotal,2) & "</b></td>"
				'response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dPenSubTotal,2) & "</b></td>"
				'response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dCertSubTotal,2) & "</b></td>"
				response.write "</tr>"
				dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
				iTotalUnits = iTotalUnits + iSubTotalUnits
				dZoneTotal = dZoneTotal + dZoneSubTotal
				dBBSTotal = dBBSTotal + dBBSSubTotal
				'dPenTotal = dPenTotal + dPenSubTotal
				'dCertTotal = dCertTotal + dCertSubTotal
				dPermitTotal = dPermitTotal + dPermitSubTotal
				sClass = " class=""reportgrouprow"" "
			End If 
			iRowCount = 1
			sReportGroup = oRs("reportgroup")
			response.write vbcrlf & "<tr" & sClass & " height=""26""><td colspan=""9""><strong>" & sReportGroup & "</strong></td></tr>"
			iSubTotalUnits = CLng(0)
			dCostEstimateSubTotal = CDbl(0.00)
			dZoneSubTotal = CDbl(0.00)
			dBBSSubTotal = CDbl(0.00)
			'dPenSubTotal = CDbl(0.00)
			'dCertSubTotal = CDbl(0.00)
			dPermitSubTotal = CDbl(0.00)
		End If 

		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr"
		If iRowCount Mod 2 = 0 Then
			response.write " class=""altrow"""
		End If 
		response.write " height=""20"">"
		
		response.write "<td align=""center"">&nbsp;" & GetPermitNumber( oRs("permitid") ) & "</td>"
		response.write "<td align=""center"">" & FormatDateTime(oRs("issueddate"),2) & "</td>"
		response.write "<td width=""400"" nowrap>" & oRs("descriptionofwork") & "</td>"
		response.write "<td width=""400"" nowrap>" & oRs("permitaddress") & "</td>"
		response.write "<td align=""center"">" & oRs("residentialunits") & "</td>"
		iSubTotalUnits = iSubTotalUnits + CLng(oRs("residentialunits"))
		
		response.write "<td align=""right"">&nbsp;" '[" & bUsePermitJobValue & "]"
		If bUsePermitJobValue Then
			response.write FormatNumber(oRs("jobvalue"),2)
			dCostEstimateSubTotal = dCostEstimateSubTotal + CDbl(oRs("jobvalue"))
		Else 
			response.write GetCostEstimate( oRs("permitid"), oRs("isold"), dCostEstimateSubTotal, sEndDate, sStartDate ) 
		End If 
		response.write "</td>"
		response.write "<td align=""right"">&nbsp;" & GetPermitFees( oRs("permitid"), oRs("isold"), dPermitSubTotal, sEndDate, sStartDate ) & "</td>"
		response.write "<td align=""right"">&nbsp;" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "iszone", dZoneSubTotal, sEndDate, sStartDate ) & "</td>"
		response.write "<td align=""right"">&nbsp;" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "isbbs", dBBSSubTotal, sEndDate, sStartDate ) & "</td>"
		'response.write "<td align=""right"">&nbsp;" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "isrevisionpenalty", dPenSubTotal ) & "</td>"
		'response.write "<td align=""right"">&nbsp;" & GetPermitReportingFees( oRs("permitid"), oRs("isold"), "iscertofoccupancy", dCertSubTotal ) & "</td>"
		response.write "</tr>"
		response.flush
		oRs.MoveNext
	Loop
	' Print out the last subTotalLine
	response.write vbcrlf & "<tr height=""20""><td colspan=""4"">&nbsp;</td>"
	response.write "<td align=""center""><b>" & iSubTotalUnits & "</td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dCostEstimateSubTotal,2) & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dPermitSubTotal,2) & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dZoneSubTotal,2) & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dBBSSubTotal,2) & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dPenSubTotal,2) & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dCertSubTotal,2) & "</b></td>"
	response.write "</tr>"
	response.flush
	dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
	iTotalUnits = iTotalUnits + iSubTotalUnits
	dZoneTotal = dZoneTotal + dZoneSubTotal
	dBBSTotal = dBBSTotal + dBBSSubTotal
	'dPenTotal = dPenTotal + dPenSubTotal
	'dCertTotal = dCertTotal + dCertSubTotal
	dPermitTotal = dPermitTotal + dPermitSubTotal

	' Print out the TotalLine
	response.write vbcrlf & "<tr height=""26""><td colspan=""4""><b>Monthly Totals</b></td>"
	response.write "<td align=""center""><b>" & iTotalUnits & "</td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dCostEstimateTotal,2) & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dPermitTotal,2) & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dZoneTotal,2) & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dBBSTotal,2) & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dPenTotal,2) & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & FormatNumber(dCertTotal,2) & "</b></td>"
	response.write "</tr>"
	response.flush

	' Get YTDs and calculate the Previous Totals
	iYTDUnits = GetYTDResidentialUnits( sYearStart, sYearEnd, iInclude )
	iPreviousUnits = CLng(iYTDUnits) - CLng(iTotalUnits)

	If bUsePermitJobValue Then
		dYTDCostEstimate = GetYTDJobValues( sYearStart, sYearEnd, iInclude )
	Else 
		dYTDCostEstimate = GetYTDCostEstimate( sYearStart, sYearEnd, iInclude )
	End If 
	dPreviousCostEstimate = FormatNumber(CDbl(dYTDCostEstimate) - CDbl(dCostEstimateTotal),2)

	dYTDPermitFees = GetYTDPermitFees( sYearStart, sYearEnd, iInclude )
	dPreviousPermitFees = FormatNumber(CDbl(dYTDPermitFees) - CDbl(dPermitTotal),2)

	dYTDZone = GetYTDPermitReportingFees( sYearStart, sYearEnd, "iszone", iInclude )
	dPreviousZone = FormatNumber(CDbl(dYTDZone) - CDbl(dZoneTotal),2)

	dYTDBBS = GetYTDPermitReportingFees( sYearStart, sYearEnd, "isbbs", iInclude )
	dPreviousBBS = FormatNumber(CDbl(dYTDBBS) - CDbl(dBBSTotal),2)

	'dYTDPen = GetYTDPermitReportingFees( sYearStart, sYearEnd, "isrevisionpenalty", iInclude )
	'dPreviousPen = FormatNumber(CDbl(dYTDPen) - CDbl(dPenTotal),2)

	'dYTDCert = GetYTDPermitReportingFees( sYearStart, sYearEnd, "iscertofoccupancy", iInclude )
	'dPreviousCert = FormatNumber(CDbl(dYTDCert) - CDbl(dPreviousCert),2)

	' Print out the Previous Totals Line
	response.write vbcrlf & "<tr height=""26""><td colspan=""4""><b>Previous Totals</b></td>"
	response.write "<td align=""center""><b>" & iPreviousUnits & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dPreviousCostEstimate & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dPreviousPermitFees & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dPreviousZone & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dPreviousBBS & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & dPreviousPen & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & dPreviousCert & "</b></td>"
	response.write "</tr>"
	response.flush

	' Print out the YTD Line
	response.write vbcrlf & "<tr height=""26""><td colspan=""4""><b>Year To Date</b></td>"
	response.write "<td align=""center""><b>" & iYTDUnits & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dYTDCostEstimate & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dYTDPermitFees & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dYTDZone & "</b></td>"
	response.write "<td align=""right""><b>&nbsp;" & dYTDBBS & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & dYTDPen & "</b></td>"
	'response.write "<td align=""right""><b>&nbsp;" & dYTDCert & "</b></td>"
	response.write "</tr>"
	response.flush

	response.write vbcrlf & "</table></body></html>"
	response.flush
End If 

oRs.Close
Set oRs = Nothing 

%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
