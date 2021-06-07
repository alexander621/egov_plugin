<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: AVAILABILITY_SAVE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: COPYRIGHT 2006 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  SAVE FACILITY RATE AND ADD NEW RATES
'
' MODIFICATION HISTORY
' 1.0   01/17/06	JOHN STULLENBERGER - INITIAL VERSION
' 1.0   01/18/06	STEVE LOAR - CODE ADDED
' 2.0	01/22/07	JOHN STULLENBERGER - NEW VERSION WITH DIFFERENT PRICING LEVELS
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------



' CALL SUB TO SAVE TIMEPART
Call subSaveTimePart(request("iFacilityId"), request("facilitytimepartid"), request("weekday"), request("beginhour"), request("beginampm"), request("endhour"), request("endampm"), request("description"), request("rate"),request("ratechoice"))




'------------------------------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------------------------------


'-------------------------------------------------------------------------------------------------------------------------------------
' SUB SUBSAVETIMEPART(IFACILITYID, IFACILITYTIMEPARTID, IWEEKDAY, SBEGINHOUR, SBEGINAMPM, SENDHOUR, SENDAMPM, SDESCRIPTION, IRATE)
'-------------------------------------------------------------------------------------------------------------------------------------
Sub subSaveTimePart( iFacilityId, iFacilitytimepartid, iWeekday, sBeginhour, sBeginampm, sEndhour, sEndampm, sDescription, iRate,iRateID)

	If iFacilitytimepartid = "0" Then
		' INSERT NEW RECORDS
		sSql = "INSERT INTO egov_facilitytimepart (orgid, facilityid, beginhour, beginampm, endhour, endampm, weekday, description, rateid) Values (" 
		sSql = sSQL & Session("OrgID") & ", " & iFacilityId & ", '" & sBeginhour & "', '" & sBeginampm & "', '" & sEndhour & "', '" & sEndampm & "', " & iWeekday & ", '" & sDescription & "','" & irateid & "')"
	Else 
		' UPDATE EXISTING RECORDS
		sSQL = "UPDATE egov_facilitytimepart SET beginhour = '" & sBeginhour 
		sSQL = sSQL & "', beginampm = '" & sBeginampm & "', endhour = '" & sEndhour & "', endampm = '" & sEndampm & "', weekday = " & iWeekday 
		sSQL = sSQL & " , description = '" & sDescription & "', rateid='" & iRateID & "', rate = NULL "
		sSQL = sSQL & " WHERE facilitytimepartid = " & iFacilitytimepartid & ""
	End If

	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO facility availability page
	response.redirect( "facility_availability.asp?facilityid=" & iFacilityId )

End Sub
%>