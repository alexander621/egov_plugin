<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: recreation_details_update.asp
' AUTHOR: Steve Loar
' CREATED: 3/10/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This code saves changes to the Facility Field Values.  
'				It is called from facility_reservation_edit.asp
'
' MODIFICATION HISTORY
' 1.0   03/10/2006 Steve Loar - Initial Version Created
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sVariable_name, sVariable_value, iPaymentId

iPaymentId = CLng(request("facilityscheduleid"))

' Loop through the request object and do the update 
' Each request object is a pair of the facilityvalueid and field value
For Each sVariable_name In Request.Form
	sVariable_value = request.form(sVariable_name)
	'response.write sVariable_name & " - " & sVariable_value & "<br />"
	If Left(sVariable_name, 6) = "field_" Then
		UpdateDetails iPaymentId, sVariable_name, sVariable_value
	End If
Next 

' UPDATE INTERNAL NOTE
UpdateNote iPaymentId, request("internalnote")


' Take them back to the facility_reservation_edit page
if Session("RedirectPage") <> "" then
	response.redirect( Session("RedirectPage") )
else
	response.write "Your data was updated successfully"
	response.end
end if


'------------------------------------------------------------------------------------------------------------
' FUNCTION UpdateDetails( iPaymentId, iFacilityValueId, sFieldValue )
'------------------------------------------------------------------------------------------------------------
Sub UpdateDetails( ByVal iPaymentId, ByVal iFacilityValueId, ByVal sFieldValue )
	Dim sSql
			
	sFieldValue = DBsafe(sFieldValue)

	If request("operation") = "update" Then
		sSql = "UPDATE egov_facility_field_values SET fieldvalue = '" & sFieldValue & "' "
		sSql = sSql & "WHERE facilityvalueId = " & Replace(iFacilityValueId,"field_","") 
		sSql = sSql & " AND paymentid = " & iPaymentId
	Else
		sSql = "INSERT INTO egov_facility_field_values ( fieldid, fieldvalue, paymentid ) VALUES ( " & Replace(iFacilityValueId,"field_","") & ", '" & sFieldValue & "', " & iPaymentId & " )"
	End If

	'response.write sSql & "<br><br>"

	'response.End 

	RunSQLStatement sSql

End Sub


'------------------------------------------------------------------------------------------------------------
' FUNCTION DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )

	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

	DBsafe = Replace( strDB, "'", "''" )

End Function


'------------------------------------------------------------------------------------------------------------
' SUB UPDATENOTE(IFACILITYVALUEID, SNOTE)
'------------------------------------------------------------------------------------------------------------
Sub UpdateNote( ByVal iFacilityValueId, ByVal sNote )
	Dim sSql
	
	sSql = "UPDATE egov_facilityschedule SET internalnote = '" & dbsafe(sNote) & "' WHERE facilityscheduleid = " & iFacilityValueId

	RunSQLStatement sSql

End Sub


%>
