<%

Call subDeleteWaiver(request("iWaiverId"), request("iFacilityId"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subDeleteWaiver(iWaiverId, iFacilityId)
' AUTHOR: Steve Loar
' CREATED: 01/22/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subDeleteWaiver(iWaiverId, iFacilityId)
	
	sSql = "Delete From egov_facilitywaivers Where waiverid = " &  iWaiverId & ""
'	response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
		sSql = "Delete from egov_waivers where waiverid = " & iWaiverId & ""
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO facility waivers page
	response.redirect( "facility_waivers.asp?facilityid=" & iFacilityId )

End Sub
%>