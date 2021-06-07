<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: monthlypermitreportexport.asp
' AUTHOR: Steve Loar
' CREATED: 07/20/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permits issued, dumped to excel - for all clients except Loveland, OH
'
' MODIFICATION HISTORY
' 1.0   07/20/2009	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iStartMonth, iStartYear, sStartDate, sYearStart, iEndMonth, iEndYear, sEndDate, sYearEnd
Dim sSql, oRs, iRowCount, sReportGroup, dCostEstimateSubTotal, dCostEstimateTotal, sClass
Dim iSubTotalUnits, iTotalUnits, dZoneSubTotal, dZoneTotal, dBBSSubTotal, dBBSTotal, sSearch
Dim dPenSubTotal, dPenTotal, dCertSubTotal, dCertTotal, dPermitSubTotal, dPermitTotal
Dim iYTDUnits, iPreviousUnits, dYTDCostEstimate, dPreviousCostEstimate, dYTDPermitFees
Dim dPreviousPermitFees, dYTDZone, dPreviousZone, dYTDBBS, dPreviousBBS, dYTDPen, iPermitCategoryId
Dim dPreviousPen, dYTDCert, dPreviousCert, iInclude, iIncludeJobValue, bUsePermitJobValue

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
Response.AddHeader "Content-Disposition", "attachment;filename=Monthly_Permit_Report_" & sDate & ".xls"

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
	iIncludeJobValue = 1
	bUsePermitJobValue = False 
End If

If request("permitcategoryid") <> "" Then
	If CLng(iPermitCategoryId) > CLng(0) Then
		sSearch = " AND P.permitcategoryid = " & iPermitCategoryId
	Else 
		sSearch = ""
	End If 
Else 
	sSearch = ""
End If 




'sSql = "SELECT P.permitid, P.issueddate, P.jobvalue, ISNULL(P.residentialunits,0) AS residentialunits, "
'sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 0 AS isold, U.reportgroup, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
'sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, O.occupancytype  "
'sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitusetypes U, egov_permitlocationrequirements R, egov_occupancytypes O "
'sSql = sSql & " WHERE P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "' AND P.permitlocationrequirementid = R.permitlocationrequirementid "
'sSql = sSql & " AND (P.occupancytypeid IS NULL OR P.occupancytypeid = O.occupancytypeid) "
'sSql = sSql & " AND A.permitid = P.permitid AND P.usetypeid = U.usetypeid AND P.orgid = " & session("orgid")
'sSql = sSql & sIsVoided & sSearch
'sSql = sSql & " UNION ALL "
'sSql = sSql & " SELECT P.permitid, P.issueddate, P.jobvalue, 0 AS residentialunits, "
'sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 1 AS isold, U.reportgroup, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
'sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, O.occupancytype  "
'sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitinvoices I, egov_permitusetypes U, egov_permitlocationrequirements R, egov_occupancytypes O "
'sSql = sSql & " WHERE A.permitid = P.permitid AND I.permitid = P.permitid AND I.invoicedate > P.issueddate AND P.issueddate < '" & sStartDate & "' "
'sSql = sSql & " AND (P.occupancytypeid IS NULL OR P.occupancytypeid = O.occupancytypeid) "
'sSql = sSql & " AND I.invoicedate >= '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' AND P.permitlocationrequirementid = R.permitlocationrequirementid AND P.usetypeid = U.usetypeid AND P.orgid = " & session("orgid")
'sSql = sSql & " AND I.isvoided = 0 AND I.allfeeswaived = 0 " & sIsVoided & sSearch
'sSql = sSql & " ORDER BY U.reportgroup, P.issueddate, P.permitid, isold"
	sSql = "SELECT P.permitid, P.issueddate, P.jobvalue, P.isvoided, ISNULL(P.residentialunits,0) AS residentialunits, "
	sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 0 AS isold, U.reportgroup, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, O.occupancytype "
	sSql = sSql & " FROM egov_permits P  "
	sSql = sSql & " INNER JOIN egov_permitlocationrequirements R ON P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " INNER JOIN egov_permitaddress A ON A.permitid = P.permitid "
	sSql = sSql & " INNER JOIN egov_permitusetypes U ON P.usetypeid = U.usetypeid "
	sSql = sSql & " LEFT JOIN egov_occupancytypes O ON P.occupancytypeid = O.occupancytypeid "
	sSql = sSql & " WHERE P.issueddate >= '" & sStartDate & "' AND P.issueddate < '" & sEndDate & "' "
	sSql = sSql & " AND P.orgid = " & session("orgid")
	sSql = sSql & sIsVoided & sSearch
	sSql = sSql & " UNION ALL "
	sSql = sSql & " SELECT DISTINCT P.permitid, P.issueddate, P.jobvalue, P.isvoided, 0 AS residentialunits, "
	sSql = sSql & " ISNULL(P.descriptionofwork,'') AS descriptionofwork, 1 AS isold, U.reportgroup, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress, O.occupancytype "
	sSql = sSql & " FROM egov_permits P  "
	sSql = sSql & " INNER JOIN egov_permitlocationrequirements R ON P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " INNER JOIN egov_permitusetypes U ON P.usetypeid = U.usetypeid "
	sSql = sSql & " INNER JOIN egov_permitaddress A ON A.permitid = P.permitid "
	sSql = sSql & " INNER JOIN egov_permitinvoices I ON I.permitid = P.permitid "
	sSql = sSql & " LEFT JOIN egov_occupancytypes O ON P.occupancytypeid = O.occupancytypeid "
	sSql = sSql & " WHERE I.invoicedate > P.issueddate AND P.issueddate < '" & sStartDate & "' "
	sSql = sSql & " AND I.invoicedate >= '" & sStartDate & "' AND I.invoicedate < '" & sEndDate & "' AND P.orgid = " & session("orgid")
	sSql = sSql & " AND I.isvoided = 0 AND I.allfeeswaived = 0 " & sIsVoided & sSearch
	sSql = sSql & " ORDER BY U.reportgroup, P.issueddate, P.permitid, isold"
'response.write sSql & "<br /><br />"

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSQL, Application("DSN"), 3, 1

If Not oRs.EOF Then 
	response.write vbcrlf & "<html>"
	response.write vbcrlf & "<style>  "
	response.write " .moneystyle "
	response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
	response.write vbcrlf & "</style>"
	response.write vbcrlf & "<body>"
	response.write vbcrlf & "<table cellpadding=""4"" cellspacing=""0"" border=""1"">"
	response.write vbcrlf & "<tr height=""30""><th>Permit #</th><th>Issued Date</th><th>Description of Work</th><th>Address/Location</th><th>Occupancy<br />Type</th><th>Residential Units</th><th>Est. Cost</th>"
	response.write "<th>Permit Fees</th></tr>"   
	response.flush
	Do While Not oRs.EOF
		If sReportGroup <> oRs("reportgroup") Then
			If sReportGroup <> "None" Then
				' Print out a subTotalLine
				response.write vbcrlf & "<tr height=""20""><td colspan=""5"">&nbsp;</td>"
				response.write "<td align=""center""><b>" & iSubTotalUnits & "</b></td>"
				response.write "<td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dCostEstimateSubTotal),2,,,0) & "</b></td>"
				response.write "<td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dPermitSubTotal),2,,,0) & "</b></td>"
				response.write "</tr>"
				dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
				iTotalUnits = iTotalUnits + iSubTotalUnits
				dPermitTotal = dPermitTotal + dPermitSubTotal
				sClass = " class=""reportgrouprow"" "
			End If 
			iRowCount = 1
			sReportGroup = oRs("reportgroup")
			response.write vbcrlf & "<tr" & sClass & " height=""26""><td colspan=""9""><strong>" & sReportGroup & "</strong></td></tr>"
			iSubTotalUnits = CLng(0)
			dCostEstimateSubTotal = CDbl(0.00)
			dPermitSubTotal = CDbl(0.00)
		End If 

		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr"
		If iRowCount Mod 2 = 0 Then
			response.write " class=""altrow"""
		End If 
		response.write " height=""20"">"
		
		response.write "<td align=""center"">&nbsp;" & GetPermitNumber( oRs("permitid") ) & "</td>"
		response.write "<td align=""center"">" & DateValue(oRs("issueddate")) & "</td>"
		response.write "<td width=""400"" nowrap>" & oRs("descriptionofwork") & "</td>"

		'response.write "<td width=""400"" nowrap>" & oRs("permitaddress") & "</td>"
		response.write "<td width=""400"" nowrap=""nowrap"">"
		Select Case oRs("locationtype")
			Case "address"
				response.write oRs("permitaddress")

			Case "location"
				response.write oRs("permitlocation")

			Case Else
				response.write ""

		End Select  
		response.write "</td>"

		response.write "<td align=""center"">" & oRs("occupancytype") & "</td>"
		response.write "<td align=""center"">" & oRs("residentialunits") & "</td>"
		iSubTotalUnits = iSubTotalUnits + CLng(oRs("residentialunits"))
		
		'response.write "<td align=""right"">&nbsp;" & GetCostEstimate( oRs("permitid"), oRs("isold"), dCostEstimateSubTotal, sEndDate, sStartDate ) & "</td>"
		response.write "<td align=""right"" class=""moneystyle"">" 
		If bUsePermitJobValue Then
			response.write FormatNumber(CDbl(oRs("jobvalue")),2,,,0)
			dCostEstimateSubTotal = dCostEstimateSubTotal + CDbl(oRs("jobvalue"))
		Else 
			response.write FormatNumber(CDbl(GetCostEstimate( oRs("permitid"), oRs("isold"), dCostEstimateSubTotal, sEndDate, sStartDate )),2,,,0)
		End If 
		'response.write "<!-- isold: " & oRs("isold") & ", dCostEstimateSubTotal: " & dCostEstimateSubTotal & ", sEndDate: " & sEndDate & ", sStartDate: " & sStartDate & " -->"
		response.write "</td>"
		response.write "<td align=""right"" class=""moneystyle"">" & FormatNumber(CDbl(GetAllPermitFees( oRs("permitid"), oRs("isold"), dPermitSubTotal, sEndDate, sStartDate )),2,,,0) & "</td>"
		response.write "</tr>"
		response.flush
		oRs.MoveNext
	Loop
	' Print out the last subTotalLine
	response.write vbcrlf & "<tr height=""20""><td colspan=""5"">&nbsp;</td>"
	response.write "<td align=""center""><b>" & iSubTotalUnits & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dCostEstimateSubTotal),2,,,0) & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dPermitSubTotal),2,,,0) & "</b></td>"
	response.write "</tr>"
	response.flush
	dCostEstimateTotal = dCostEstimateTotal + dCostEstimateSubTotal
	iTotalUnits = iTotalUnits + iSubTotalUnits
	dPermitTotal = dPermitTotal + dPermitSubTotal

	' Print out the TotalLine
	response.write vbcrlf & "<tr height=""26""><td colspan=""5""><b>Monthly Totals</b></td>"
	response.write "<td align=""center"" class=""intstyle""><b>" & iTotalUnits & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dCostEstimateTotal),2,,,0) & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dPermitTotal),2,,,0) & "</b></td>"
	response.write "</tr>"
	response.flush

	' Get YTDs and calculate the Previous Totals
	iYTDUnits = GetYTDResidentialUnits( sYearStart, sYearEnd, iInclude )
	iPreviousUnits = CLng(iYTDUnits) - CLng(iTotalUnits)

	'dYTDCostEstimate = GetYTDCostEstimate( sYearStart, sYearEnd, iInclude )
	If bUsePermitJobValue Then
		dYTDCostEstimate = FormatNumber(CDbl(GetYTDJobValues( sYearStart, sYearEnd, iInclude )),2,,,0)
	Else 
		dYTDCostEstimate = FormatNumber(CDbl(GetYTDCostEstimate( sYearStart, sYearEnd, iInclude )),2,,,0)
	End If 
	dPreviousCostEstimate = FormatNumber(CDbl(dYTDCostEstimate) - CDbl(dCostEstimateTotal),2,,,0)

	dYTDPermitFees = GetAllYTDPermitFees( sYearStart, sYearEnd, iInclude )
	dPreviousPermitFees = FormatNumber(CDbl(dYTDPermitFees) - CDbl(dPermitTotal),2,,,0)

	' Print out the Previous Totals Line
	response.write vbcrlf & "<tr height=""26""><td colspan=""5""><b>Previous Totals</b></td>"
	response.write "<td align=""center""><b>" & iPreviousUnits & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & dPreviousCostEstimate & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & dPreviousPermitFees & "</b></td>"
	response.write "</tr>"
	response.flush

	' Print out the YTD Line
	response.write vbcrlf & "<tr height=""26""><td colspan=""5""><b>Year To Date</b></td>"
	response.write "<td align=""center""><b>" & iYTDUnits & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & dYTDCostEstimate & "</b></td>"
	response.write "<td align=""right"" class=""moneystyle""><b>" & dYTDPermitFees & "</b></td>"
	response.write "</tr>"
	response.flush

	response.write vbcrlf & "</table></body></html>"
End If 

oRs.Close
Set oRs = Nothing 

%>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
