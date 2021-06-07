<%

Call subSavePOC( request("pocId"), request("sName"), request("sEmail"), request("sPhone") )


'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' SUB subSaveInstructor( iPocId, sName, sEmail, sPhone )
' AUTHOR: Steve Loar
' CREATED: 05/10/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'--------------------------------------------------------------------------------------------------
Sub subSavePOC( iPocId, sName, sEmail, sPhone )
	Dim sSql, oCmd

	sName = DBsafe( sName )
	sEmail = DBsafe( sEmail )
	sPhone = DBsafe( sPhone )

	If clng(iPocId) = 0 Then
		' Insert new records
		sSql = "INSERT INTO egov_class_pointofcontact ( orgid, name, email, phone ) Values ( " 
		sSql = sSql & Session("OrgID") & ", '" & sName & "', '" & sEmail & "', '" & sPhone & "' )" 
	Else 
		' Update existing records
		sSQL = "UPDATE egov_class_pointofcontact SET name = '" & sName & "', email= '" & sEmail & "', phone= '" & sPhone & "' " 
		sSql = sSql & " WHERE pocid = " & iPocId 
	End If

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
		.CommandText = sSql
		.Execute
	End With
	Set oCmd = Nothing

	' REDIRECT TO point of poc management page
	response.redirect "poc_mgmt.asp"

End Sub


%>

<!-- #include file="../includes/common.asp" //-->
