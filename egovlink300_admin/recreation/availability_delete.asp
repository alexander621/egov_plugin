<%
Call subDeleteTimePart(request("iFacilitytimepartid"), request("iFacilityId"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subDeleteTimePart(iFacilitytimepartid, iFacilityID)
' AUTHOR: Steve Loar
' CREATED: 01/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subDeleteTimePart(iFacilitytimepartid, iFacilityID)
	
	' Delete from the facilitytimepart table
	sSQL = "DELETE FROM egov_facilitytimepart WHERE Facilitytimepartid =" & iFacilitytimepartid  & ""

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO Facility availability PAGE
	response.redirect( "facility_availability.asp?facilityid=" & iFacilityId )

End Sub
%>