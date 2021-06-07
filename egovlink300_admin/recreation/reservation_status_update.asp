<%
Call subUpdateStatus(request("applies"), request("selstatus"), request("ireservationid"), request("ioccurrenceid"), request("sCancelReason"))


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB SUBDELETETERM(ITERMID, IFACILITYID)
'--------------------------------------------------------------------------------------------------
Sub subUpdateStatus(sApplies,sStatus,iReservationID,ioccurenceid,scanceldescription)
	
	' DETERMINE IF WE UPDATE THE OCCURENCE OR THE SERIES
	If UCASE(sApplies) = "ALL" Then
		' UPDATE THE SERIES
		sSQL = "UPDATE egov_facilityschedule set datecancelled='" & Date() & "', status = '" & request("selstatus") & "', canceldescription='" & scanceldescription & "' WHERE facilityrecurrenceid = '" & ioccurenceid & "'"
	Else
		' UPDATE THE SINGLE RESERVATION
		sSQL = "UPDATE egov_facilityschedule set datecancelled='" & Date() & "',status = '" & request("selstatus") & "', canceldescription='" & DBsafe(scanceldescription) & "' WHERE facilityscheduleid = '" & iReservationID & "'"
	End If

	' NEED TO ADD CODE TO CHECK FOR EXISTING RESERVATION IF CHANGING STATUS TO RESERVED

	' UPDATE RESERVATION
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO Facility rate PAGE
	response.redirect( "facility_reservation_edit.asp?ireservationid=" & ireservationid )

End Sub

'----------------------------------------
'  Make buffer Database 'safe'
'  Useful in building SQL Strings
'    strSQL="SELECT *....WHERE Value='" & DBSafe(strValue) & "';"
'----------------------------------------
Function DBsafe( strDB )
  If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function
  DBsafe = Replace( strDB, "'", "''" )
End Function

%>