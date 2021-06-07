<%
Call subDeleteRate(request("iRateId"), request("iFacilityId"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subDeleteRate(iRateId, iFacilityID)
' AUTHOR: Steve Loar
' CREATED: 01/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subDeleteRate(iRateId, iFacilityID)
	
	' Delete from the rate table
	sSQL = "DELETE FROM egov_rate WHERE rateid =" & iRateId  & ""

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO Facility rate PAGE
	response.redirect( "facility_rates.asp?facilityid=" & iFacilityId )

End Sub
%>