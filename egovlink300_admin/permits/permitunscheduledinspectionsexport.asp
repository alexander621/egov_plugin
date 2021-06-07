<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitunscheduledinspectionsexport.asp
' AUTHOR: Steve Loar
' CREATED: 05/18/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Report of permit inspections that have not been completed, dumped to excel
'
' MODIFICATION HISTORY
' 1.0   05/18/2010	Steve Loar - INITIAL VERSION
' 1.1	11/15/2010	Steve Loar - Added permit category
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' SET UP PAGE OPTIONS
sDate = Right("0" & Month(Date()),2) & Right("0" & Day(Date()),2) & Year(Date())
server.scripttimeout = 9000
Response.ContentType = "application/vnd.ms-excel"
Response.AddHeader "Content-Disposition", "attachment;filename=UnscheduledInspectionReport_" & sDate & ".xls"

Dim sSearch

'	sSearch = session("sSql")

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

If request("permitinspectiontypeid") <> "" Then
	iPermitInspectionTypeId = CLng(request("permitinspectiontypeid"))
	If iPermitInspectionTypeId > CLng(0) Then
		sSearch = sSearch & " AND I.permitinspectiontypeid = " & iPermitInspectionTypeId
	End If 
End If 

If request("inspectoruserid") <> "" Then 
	iInspectorUserId = CLng(request("inspectoruserid"))
	If iInspectorUserId > CLng(0) Then 
		sSearch = sSearch & " AND I.inspectoruserid = " & iInspectorUserId
	End If 
End If 

If request("permitcategoryid") <> "" Then
	If CLng(request("permitcategoryid")) > CLng(0) Then 
		sSearch = sSearch & " AND P.permitcategoryid = " & CLng(request("permitcategoryid"))
	End If 
End If 

If request("permitlocation") <> "" Then
	sSearch = sSearch & " AND P.permitlocation LIKE '%" & dbsafe(request("permitlocation")) & "%' "
End If 


DisplayInspections sSearch



'--------------------------------------------------------------------------------------------------
'SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' void DisplayInspections sSearch 
'--------------------------------------------------------------------------------------------------
Sub DisplayInspections( ByVal sSearch )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT I.permitid, I.permitinspectiontype, S.inspectionstatus, I.scheduleddate, I.inspecteddate, "
	sSql = sSql & " ISNULL(I.inspectoruserid,0) AS inspectoruserid, ISNULL(P.permitlocation,'') AS permitlocation, R.locationtype, "
	sSql = sSql & " A.residentstreetnumber, ISNULL(A.residentstreetprefix,'') AS residentstreetprefix, A.residentstreetname, ISNULL(A.streetsuffix,'') AS streetsuffix, ISNULL(A.streetdirection,'') AS streetdirection, ISNULL(A.residentunit,'') AS residentunit, "
	sSql = sSql & " dbo.fn_buildAddress(A.residentstreetnumber, A.residentstreetprefix, A.residentstreetname, A.streetsuffix, A.streetdirection ) AS permitaddress "
	sSql = sSql & " FROM egov_permitinspections I, egov_inspectionstatuses S, egov_permitaddress A, egov_permits P, egov_permitlocationrequirements R "
	sSql = sSql & " WHERE I.orgid = " & session("orgid") & " AND I.inspecteddate IS NULL AND S.isneedsinspection = 1 "
	sSql = sSql & " AND S.inspectionstatusid = I.inspectionstatusid AND P.isvoided = 0 AND P.permitlocationrequirementid = R.permitlocationrequirementid AND "
	sSql = sSql & " P.permitid = I.permitid AND A.permitid = I.permitid " & sSearch
	sSql = sSql & " ORDER BY P.permitnumberyear, P.permitnumber, P.permitid, I.inspectionorder"

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
		response.write "<th>Permit #</th><th>Inspection Type</th><th>Status</th><th>Scheduled Date</th><th>Inspector</th></tr>"
		response.flush

		Do While Not oRs.EOF
			response.write vbcrlf & "<tr height=""26"">"
			
			' Address broken out
			response.write "<td align=""center"">&nbsp;" & oRs("residentstreetnumber") & "</td>"
			response.write "<td align=""center"">&nbsp;" & oRs("residentstreetprefix") & "</td>"
			response.write "<td width=""300"" nowrap>" & oRs("residentstreetname") & "</td>"
			response.write "<td align=""center"">&nbsp;" & oRs("streetsuffix") & "</td>"
			response.write "<td align=""center"">&nbsp;" & oRs("streetdirection") & "</td>"
			response.write "<td align=""center"">&nbsp;" & oRs("residentunit") & "</td>"

			' Location
			response.write "<td align=""center"">" & Replace(oRs("permitlocation"),Chr(10),"<br />") & "</td>"

			response.write "<td align=""center"">&nbsp;" & GetPermitNumber( oRs("permitid") ) & "</td>"
			response.write "<td align=""center"">&nbsp;" & oRs("permitinspectiontype") & "</td>"
			response.write "<td align=""center"">" & oRs("inspectionstatus") & "</td>"
			response.write "<td align=""center"">&nbsp;" & oRs("scheduleddate") & "</td>"
			response.write "<td align=""center"" width=""200"" nowrap>&nbsp;" 
			If CLng(oRs("inspectoruserid")) > CLng(0) Then 
				response.write GetAdminName( oRs("inspectoruserid") )
			Else
				response.write ""
			End If 
			response.write "</td>"
			response.write "</tr>"
			response.flush
			oRs.MoveNext 
		Loop
		
		response.write vbcrlf & "</table></body></html>"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 



%>

<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
