<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitissuedreportexport.asp
' AUTHOR: Steve Loar
' CREATED: 09/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permits issued, dumped to excel
'
' MODIFICATION HISTORY
' 1.0   09/11/2008	Steve Loar - INITIAL VERSION
' 1.1	12/2/2009	Steve Loar - Added the County field (address grouping field) for Loveland, OH
' 1.2	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSearch, sFromIssuedDate, sToIssuedDate, sStreetNumber, sStreetName, sPermitNo
Dim iPermitTypeId, sApplicant, iPermitStatusId, sDisplayDateRange, iOrderBY, iPermitCategoryId

' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=Issued_Permit_Report_" & sDate & ".xls"

'Dim sSearch, sRptTitle

'sSearch = session("sSql")

'sRptTitle = vbcrlf & "<tr><th>Issued Permit Report</th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

'DisplayIssuedPermits sSearch, sRptTitle

' Handle inspection date range. always want some dates to limit the search
If request("toissueddate") <> "" And request("fromissueddate") <> "" Then
	sFromIssuedDate = request("fromissueddate")
	sToIssuedDate = request("toissueddate")
	sSearch = sSearch & " AND (P.issueddate >= '" & request("fromissueddate") & "' AND P.issueddate < '" & DateAdd("d",1,request("toissueddate")) & "' ) "
	sDisplayDateRange = "From: " & request("fromissueddate") & " &nbsp;To: " & request("toissueddate")
Else
	' initially set these to yesterday
	sFromIssuedDate = FormatDateTime(DateAdd("m",-1,Date),2)
	sToIssuedDate = FormatDateTime(Date,2)
	sDisplayDateRange = ""
End If 

' handle address pick
If request("residentstreetnumber") <> "" Then 
	sStreetNumber = request("residentstreetnumber")
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

If request("permittypeid") <> "" Then
	iPermitTypeId = CLng(request("permittypeid"))
	If iPermitTypeId > CLng(0) Then
		sSearch = sSearch & " AND P.permittypeid = " & iPermitTypeId
	End If 
End If 

If request("permitstatusid") <> "" Then
	iPermitStatusId = CLng(request("permitstatusid"))
	If iPermitStatusId > CLng(0) Then
	Select Case iPermitStatusId
			Case 1
				sSearch = sSearch & " AND (S.isissued = 1 OR S.isissuedback = 1) AND P.isvoided = 0 "
			Case 2
				sSearch = sSearch & " AND S.iscompletedstatus = 1 AND P.isvoided = 0 "
			Case 3
				sSearch = sSearch & " AND P.isonhold = 1 AND P.isvoided = 0 "
			Case 4
				sSearch = sSearch & " AND P.isexpired = 1 AND P.isvoided = 0 "
			Case 5
				sSearch = sSearch & " AND P.isvoided = 1 "
			Case 6
				sSearch = sSearch & " AND (S.isissued = 1 OR S.isissuedback = 1 OR S.iscompletedstatus = 1) AND P.isvoided = 0 "
		End Select 
	End If 
End If 

If request("applicant") <> "" Then 
	sApplicant = request("applicant")
	sSearch = sSearch & " AND ( C.company LIKE '%" & dbsafe(sApplicant) & "%' OR C.firstname LIKE '%" & dbsafe(sApplicant) & "%' OR C.lastname LIKE '%" & dbsafe(sApplicant) & "%' ) "
End If 

If request("permitcategoryid") <> "" Then
	If CLng(request("permitcategoryid")) > CLng(0) Then 
		sSearch = sSearch & " AND P.permitcategoryid = " & CLng(request("permitcategoryid"))
	End If 
End If 

If request("permitlocation") <> "" Then
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If 

If request("orderby") <> "" Then 
	iOrderBY = clng(request("orderby"))
Else
	iOrderBY = clng(1)
End If 

If clng(iOrderBy) = clng(1) Then 
	DisplayIssuedPermits sSearch
Else
	DisplayIssuedPermitsByType sSearch
End If 


'--------------------------------------------------------------------------------------------------
' void DisplayIssuedPermits sSearch
'--------------------------------------------------------------------------------------------------
Sub DisplayIssuedPermits( ByVal sSearch )
	Dim sSql, oRs, iRowCount, dJobTotal, dPaidTotal, bIsVoided, dPaidAmount, dJobValue

	iRowCount = 0
	dJobTotal = CDbl(0.00)
	dPaidTotal = CDbl(0.00)
	bIsVoided = false

	sSql = "SELECT P.permitid, P.issueddate, P.descriptionofwork, P.applieddate, ISNULL(P.jobvalue,0.00) AS jobvalue, ISNULL(P.totalpaid,0.00) AS totalpaid,"
	sSql = sSql & " CASE WHEN P.isexpired = 1 THEN 'Expired' ELSE CASE WHEN P.isonhold = 1 THEN 'On Hold' ELSE CASE WHEN P.isvoided = 1 THEN 'Voided' ELSE S.permitstatus END END END AS permitstatus, "
	sSql = sSql & " P.isvoided, T.permittype, T.permittypedesc, A.legaldescription, ISNULL(C.company,'') AS company, "
	sSql = sSql & " ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, ISNULL(A.county,'') AS county, "
	sSql = sSql & " ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitstatuses S, egov_permittypes T, egov_permitcontacts C, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.issueddate IS NOT NULL " & sSearch
	sSql = sSql & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND P.permittypeid = T.permittypeid "
	sSql = sSql & " AND P.permitid = C.permitid AND C.isapplicant = 1 AND P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " ORDER BY A.residentstreetname, A.streetsuffix, A.residentstreetnumber, P.permitnumberyear, P.permitnumberprefix, P.permitnumber, P.permitid"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<html>"
		response.write vbcrlf & "<style>  "
		response.write " .moneystyle "
		response.write vbcrlf & "{mso-style-parent:style0;mso-number-format:""\#\,\#\#0\.00"";} "
		response.write vbcrlf & "</style>"
		response.write "<body><table border=""1"" cellpadding=""4"">"
		response.write vbcrlf & "<tr height=""30""><th>Permit #</th><th>Current Status</th><th>Description Of Work</th><th>Applied Date</th><th>Issued Date</th><th>Permit Type</th>"
		response.write "<th>Number</th><th>Prefix</th><th>Street Name</th><th>Suffix</th><th>Direction</th><th>Unit</th><th>Location</th>"
		response.write "<th>" & GetOrgDisplayWithId( session("orgid"), GetDisplayId("address grouping field"), True ) & "</th>"
		response.write "<th>Legal Desc</th><th>Valuation</th><th>Fees Paid</th><th>Applicant</th></tr>"
		response.flush

		Do While Not oRs.EOF
			response.write vbcrlf & "<tr height=""22"">"
			twfPN = GetPermitNumber( oRs("permitid") )
			if twfPN = "" then twfPN = "&nbsp;"
			response.write "<td align=""center"">" & twfPN & "</td>"
			response.write "<td align=""center"">" & oRs("permitstatus") & "</td>"
			response.write "<td width=""300"" nowrap>" & oRs("descriptionofwork") & "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("applieddate"),2) & "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("issueddate"),2) & "</td>"
			response.write "<td align=""center"">" & oRs("permittype") & "</td>"
			
			'response.write "<td>&nbsp;" & oRs("permitaddress") & "</td>"
			response.write "<td align=""center"">" & oRs("residentstreetnumber") & "</td>"
			response.write "<td align=""center"">" & oRs("residentstreetprefix") & "</td>"
			response.write "<td width=""300"" nowrap>" & oRs("residentstreetname") & "</td>"
			response.write "<td align=""center"">" & oRs("streetsuffix") & "</td>"
			response.write "<td align=""center"">" & oRs("streetdirection") & "</td>"
			response.write "<td align=""center"">" & oRs("residentunit") & "</td>"

			' Location
			response.write "<td align=""center"">" & oRs("permitlocation") & "</td>"

			' County field (address grouping field)
			response.write "<td align=""center"">" & oRs("county") & "</td>"

			response.write "<td width=""600"" nowrap>&nbsp;" & oRs("legaldescription") & "</td>"
			
			bIsVoided = oRs("isvoided")
			If bIsVoided Then
				dJobValue = FormatNumber("0.00",2,,,0)
				dPaidAmount = FormatNumber("0.00",2,,,0)
			Else 
				dJobTotal = dJobTotal + CDbl(oRs("jobvalue"))
				dJobValue = FormatNumber(CDbl(oRs("jobvalue")),2,,,0)
				dPaidAmount = FormatNumber(CDbl(oRs("totalpaid")),2,,,0)
				dPaidTotal = dPaidTotal + CDbl(oRs("totalpaid"))
			End If 
			response.write "<td align=""right"" class=""moneystyle"">" & dJobValue & "</td>"
			response.write "<td align=""right"" class=""moneystyle"">" & dPaidAmount & "</td>"

			If oRs("firstname") <> "" Then 
				response.write "<td width=""300"" nowrap>" & oRs("firstname") & " " & oRs("lastname") & "</td>"
			Else
				response.write "<td width=""300"" nowrap>" & oRs("company") & "</td>"
			End If 
			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop
		' Totals row
		response.write vbcrlf & "<tr height=""30""><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>"
		response.write "<td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dJobTotal),2,,,0) & "</b></td><td align=""right"" class=""moneystyle""><b>" & FormatNumber(CDbl(dPaidTotal),2,,,0) & "</b></td><td></td></tr>"
		
		response.write vbcrlf & "</table></body></html>"
		response.flush
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' void DisplayIssuedPermitsByType sSearch
'--------------------------------------------------------------------------------------------------
Sub DisplayIssuedPermitsByType( ByVal sSearch )
	Dim sSql, oRs, iRowCount, dJobTotal, dPaidTotal, sOldType, dTypeJobTotal, dTypePaidTotal, bFirstType, bIsVoided, dPaidAmount, dJobValue

	iRowCount = 0
	dJobTotal = CDbl(0.00)
	dPaidTotal = CDbl(0.00)
	sOldType = "NONE"
	dTypeJobTotal = CDbl(0.00)
	dTypePaidTotal = CDbl(0.00)
	bFirstType = True 
	bIsVoided = false

	sSql = "SELECT P.permitid, P.issueddate, P.descriptionofwork, P.applieddate, ISNULL(P.jobvalue,0.00) AS jobvalue, ISNULL(P.totalpaid,0.00) AS totalpaid,"
	sSql = sSql & " CASE WHEN P.isexpired = 1 THEN 'Expired' ELSE CASE WHEN P.isonhold = 1 THEN 'On Hold' ELSE CASE WHEN P.isvoided = 1 THEN 'Voided' ELSE S.permitstatus END END END AS permitstatus, "
	sSql = sSql & " P.isvoided, T.permittype, T.permittypedesc, A.legaldescription, ISNULL(C.company,'') AS company, "
	sSql = sSql & " ISNULL(C.firstname,'') AS firstname, ISNULL(C.lastname,'') AS lastname, ISNULL(A.county,'') AS county, "
	sSql = sSql & " ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permits P, egov_permitaddress A, egov_permitstatuses S, egov_permittypes T, egov_permitcontacts C, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE P.orgid = " & session("orgid") & " AND P.issueddate IS NOT NULL " & sSearch
	sSql = sSql & " AND A.permitid = P.permitid AND P.permitstatusid = S.permitstatusid AND P.permittypeid = T.permittypeid "
	sSql = sSql & " AND P.permitid = C.permitid AND C.isapplicant = 1 AND P.permitlocationrequirementid = R.permitlocationrequirementid "
	sSql = sSql & " ORDER BY T.permittype, P.permitnumberyear, P.permitnumberprefix, P.permitnumber, P.permitid"
	'response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<html><body><table border=""1"" cellpadding=""4"">"
		response.write vbcrlf & "<tr height=""30""><th>Permit #</th><th>Current Status</th><th>Description Of Work</th><th>Applied Date</th><th>Issued Date</th><th>Permit Type</th>"
		response.write "<th>Number</th><th>Prefix</th><th>Street Name</th><th>Suffix</th><th>Direction</th><th>Unit</th><th>Location</th>"
		response.write "<th>" & GetOrgDisplayWithId( session("orgid"), GetDisplayId("address grouping field"), True ) & "</th>"
		response.write "<th>Legal Desc</th><th>Valuation</th><th>Fees Paid</th><th>Applicant</th></tr>"
		response.flush

		Do While Not oRs.EOF
			If sOldType <> oRs("permittype") Then
				If Not bFirstType Then 
					' Print out a sub total row here
					response.write vbcrlf & "<tr height=""22""><td colspan=""15"" align=""right""><b>" & sOldType & "</b></td><td align=""right""><b>&nbsp;" & FormatNumber(CDbl(dTypeJobTotal),2,,,0) & "</b></td><td align=""right""><b>&nbsp;" & FormatNumber(CDbl(dTypePaidTotal),2,,,0) & "</b></td><td></td></tr>"
				Else
					bFirstType = False 
				End If 
				sOldType = oRs("permittype")
				dTypeJobTotal = CDbl(0.00)
				dTypePaidTotal = CDbl(0.00)
				iRowCount = 0
			End If 

			response.write vbcrlf & "<tr height=""22"">"
			response.write "<td align=""center"">&nbsp;" & GetPermitNumber( oRs("permitid") ) & "</td>"
			response.write "<td align=""center"">" & oRs("permitstatus") & "</td>"
			response.write "<td width=""300"" nowrap>" & oRs("descriptionofwork") & "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("applieddate"),2) & "</td>"
			response.write "<td align=""center"">" & FormatDateTime(oRs("issueddate"),2) & "</td>"
			response.write "<td align=""center"">" & oRs("permittype") & "</td>"
			
			response.write "<td align=""center"">" & oRs("residentstreetnumber") & "</td>"
			response.write "<td align=""center"">" & oRs("residentstreetprefix") & "</td>"
			response.write "<td width=""300"" nowrap>" & oRs("residentstreetname") & "</td>"
			response.write "<td align=""center"">" & oRs("streetsuffix") & "</td>"
			response.write "<td align=""center"">" & oRs("streetdirection") & "</td>"
			response.write "<td align=""center"">" & oRs("residentunit") & "</td>"

			' Location
			response.write "<td align=""center"">" & Replace(oRs("permitlocation"),Chr(10),"<br />") & "</td>"

			' County field (address grouping field)
			response.write "<td align=""center"">" & oRs("county") & "</td>"

			response.write "<td width=""600"" nowrap>&nbsp;" & oRs("legaldescription") & "</td>"
			
			bIsVoided = oRs("isvoided")
			If bIsVoided Then
				dJobValue = FormatNumber("0.00",2,,,0)
				dPaidAmount = FormatNumber("0.00",2,,,0)
			Else 
				dJobTotal = dJobTotal + CDbl(oRs("jobvalue"))
				dJobValue = FormatNumber(CDbl(oRs("jobvalue")),2,,,0)
				dTypeJobTotal = dTypeJobTotal + CDbl(oRs("jobvalue"))
				
				dPaidAmount = FormatNumber(CDbl(oRs("totalpaid")),2,,,0)
				dPaidTotal = dPaidTotal + CDbl(oRs("totalpaid"))
				dTypePaidTotal = dTypePaidTotal + CDbl(oRs("totalpaid"))
			End If 
			response.write "<td align=""right"">&nbsp;" & dJobValue & "</td>"
			response.write "<td align=""right"">&nbsp;" & dPaidAmount & "</td>"

			If oRs("firstname") <> "" Then 
				response.write "<td width=""300"" nowrap>" & oRs("firstname") & " " & oRs("lastname") & "</td>"
			Else
				response.write "<td width=""300"" nowrap>" & oRs("company") & "</td>"
			End If 
			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop

		' Print out a sub total row here
		response.write vbcrlf & "<tr height=""22""><td colspan=""15"" align=""right""><b>" & sOldType & "</b></td><td align=""right""><b>&nbsp;" & FormatNumber(CDbl(dTypeJobTotal),2,,,0) & "</b></td><td align=""right""><b>&nbsp;" & FormatNumber(dTypePaidTotal,2,,,0) & "</b></td><td></td></tr>"

		' Totals row
		response.write vbcrlf & "<tr height=""30""><td colspan=""15"" align=""right""><b>Total</b></td>"
		response.write "<td align=""right""><b>" & FormatNumber(CDbl(dJobTotal),2,,,0) & "</b></td><td align=""right""><b>&nbsp;" & FormatNumber(CDbl(dPaidTotal),2,,,0) & "</b></td><td></td></tr>"
		
		response.write vbcrlf & "</table></body></html>"
		response.flush
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

%>

<!-- #include file="permitcommonfunctions.asp" //-->
<!-- #include file="../includes/common.asp" //-->


