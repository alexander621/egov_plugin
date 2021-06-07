<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectionreportexport.asp
' AUTHOR: Steve Loar
' CREATED: 09/05/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permit inspections dumped to excel
'
' MODIFICATION HISTORY
' 1.0   09/05/2008	Steve Loar - INITIAL VERSION
' 1.1	05/14/2010	Steve Loar - Changed query to skip the scheduled date requirement
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
	' SET UP PAGE OPTIONS
	sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
	server.scripttimeout = 9000
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment;filename=Permit_Inspection_Report_" & sDate & ".xls"

	Dim sSearch, sRptTitle

	sSearch = session("sSql")
	
	'sRptTitle = vbcrlf & "<tr height=""30""><th>Permit Inspection Report</th><th></th><th></th><th></th><th></th><th></th><th></th></tr>"

	DisplayInspections sSearch, sRptTitle

'--------------------------------------------------------------------------------------------------
'SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayInspections sSearch, sRptTitle
'--------------------------------------------------------------------------------------------------
Sub DisplayInspections( ByVal sSearch, ByVal sRptTitle )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT I.permitid, I.permitinspectiontype, S.inspectionstatus, I.scheduleddate, I.inspecteddate, "
	sSql = sSql & " U.firstname, U.lastname, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S, users U, egov_permitaddress A, egov_permits P, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE I.orgid = " & session("orgid") & " AND I.inspecteddate IS NOT NULL AND P.permitlocationrequirementid = R.permitlocationrequirementid AND "
	sSql = sSql & " S.inspectionstatusid = I.inspectionstatusid AND U.userid = I.inspectoruserid AND "
	sSql = sSql & " P.permitid = I.permitid AND A.permitid = I.permitid " & sSearch
	sSql = sSql & " ORDER BY I.inspecteddate, I.permitid"
	'response.write sSql & "<br /><br />"
	' AND I.scheduleddate IS NOT NULL 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<html><body><table border=""1"" cellpadding=""4"">"
		'response.write sRptTitle
		response.write vbcrlf & "<tr height=""30"">"
		'response.write "<th>Address</th>"
		response.write "<th>Number</th><th>Prefix</th><th>Street Name</th><th>Suffix</th><th>Direction</th><th>Unit</th><th>Location</th>"
		response.write "<th>Permit #</th><th>Inspection Type</th><th>Result</th><th>Scheduled Date</th><th>Inspected Date</th><th>Inspector</th></tr>"
		response.flush

		Do While Not oRs.EOF
			response.write vbcrlf & "<tr height=""26"">"
			
			' Address fields broken out
			response.write "<td align=""center"">" & oRs("residentstreetnumber") & "</td>"
			response.write "<td align=""center"">" & oRs("residentstreetprefix") & "</td>"
			response.write "<td width=""300"" nowrap>" & oRs("residentstreetname") & "</td>"
			response.write "<td align=""center"">" & oRs("streetsuffix") & "</td>"
			response.write "<td align=""center"">" & oRs("streetdirection") & "</td>"
			response.write "<td align=""center"">" & oRs("residentunit") & "</td>"

			' Location
			response.write "<td align=""center"">" & Replace(oRs("permitlocation"),Chr(10),"<br />") & "</td>"

			response.write "<td align=""center"">&nbsp;" & GetPermitNumber( oRs("permitid") ) & "</td>"
			response.write "<td align=""center"">" & oRs("permitinspectiontype") & "</td>"
			response.write "<td align=""center"">" & oRs("inspectionstatus") & "</td>"
			response.write "<td align=""center"">" & oRs("scheduleddate") & "</td>"
			response.write "<td align=""center"">" & oRs("inspecteddate") & "</td>"
			response.write "<td align=""center"" width=""200"" nowrap>" & oRs("firstname") & " " & oRs("lastname") & "</td>"
			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop
		
		response.write vbcrlf & "</table></body></html>"
		response.flush
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 

%>

<!-- #include file="permitcommonfunctions.asp" //-->
