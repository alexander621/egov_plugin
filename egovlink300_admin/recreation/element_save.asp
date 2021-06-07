<%

Call subSaveElement(request("elementid"), request("iFacilityId"), request("content"), request("sequence"), request("alt_tag"), request("element_type"))

'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB subSaveElement(iElementId, iFacilityId, sContent, iSequence, sAlt_tag, sElement_type)
' AUTHOR: Steve Loar
' CREATED: 01/25/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subSaveElement(iElementId, iFacilityId, sContent, iSequence, sAlt_tag, sElement_type)
	
	If iElementId = "0" Then
		' Insert new records
		sSql = "INSERT INTO egov_facilityelements (facilityid, content, sequence, alt_tag, element_type) Values (" 
		sSql = sSql & iFacilityId & ", '" & dbsafe(sContent) & "', " & isequence & ", '" & dbsafe(sAlt_tag) & "', '" & dbsafe(sElement_type) & "' )"
	Else 
		' Update existing records
		sSqL = "UPDATE egov_facilityelements SET content = '" & dbsafe(sContent) & "', sequence = " & iSequence 
		sSql = sSql & ", alt_tag = '" & dbsafe(sAlt_tag) & "', element_type = '" & dbsafe(sElement_type) & "' WHERE elementid = " & iElementId & ""
	End If
	'response.write sSQL
	
	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO facility edit page
	response.redirect( "facility_edit.asp?facilityid=" & iFacilityId )

End Sub


'------------------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then 
		sNewString = strDB
	Else 
		sNewString = Replace( strDB, "'", "''" )
		'sNewString = Replace( sNewString, "<", "&lt;" )
	End If 

	DBsafe = sNewString
End Function


%>