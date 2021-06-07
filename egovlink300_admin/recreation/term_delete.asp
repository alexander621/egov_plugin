<%
Call subDeleteTerm(request("iTermId"), request("iFacilityId"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBDELETETERM(ITERMID, IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub subDeleteTerm(iTermId, iFacilityID)
	
	sSQL = "DELETE FROM egov_recreation_terms WHERE termid =" & iTermId  & ""

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO Facility rate PAGE
	response.redirect( "facility_terms.asp?facilityid=" & iFacilityId )

End Sub
%>