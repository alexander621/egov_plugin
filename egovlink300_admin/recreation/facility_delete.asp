<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: facility_delete.asp
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
''
' DESCRIPTION:  Delete a facility.
'
' MODIFICATION HISTORY
' 1.0   01/17/06	JOHN STULLENBERGE - INITIAL VERSION
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

DeleteFacility CLng(request("ifacilityid"))

Sub DeleteFacility( ByVal iFacilityId )
	Dim sSql

	' Delete from the timepart table
	sSql = "DELETE FROM egov_facilitytimepart WHERE facilityid = " & iFacilityId 
	RunSQLStatement sSql 

	' Delete from the facility table
	sSql = "DELETE FROM egov_facility WHERE facilityid = " & iFacilityId
	RunSQLStatement sSql 

	' REDIRECT TO Facility MANANGEMENT PAGE
	response.redirect "facility_management.asp"

End Sub
%>