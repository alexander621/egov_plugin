<%

Call subSaveRate(request("iRateId"), request("iFacilityId"), request("sDescription"), request("iRateValue"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subSaveRate(iRateId, iFacilityId, sDesc, iValue)
' AUTHOR: Steve Loar
' CREATED: 01/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subSaveRate(iRateId, iFacilityId, sDesc, iValue)
	
	If iRateId = "0" Then
		' Insert new records
		sSql = "INSERT INTO egov_rate (orgid, facilityid, ratedescription, ratevalue) Values (" & Session("OrgID") & ", " & iFacilityId & ", '" & sDesc & "', " & iValue & " )"
	Else 
		' Update existing records
		sSQL = "UPDATE egov_rate SET ratedescription = '" & sDesc & "', ratevalue = " & iValue & " WHERE rateid = " & iRateId & ""
	End If
'	response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO facility rates page
	response.redirect( "facility_rates.asp?facilityid=" & iFacilityId )

End Sub
%>