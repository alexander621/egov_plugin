<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: org_display_update.asp
' AUTHOR: Steve Loar
' CREATED: 8/21/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates or updates the organizations' displays
'
' MODIFICATION HISTORY
' 1.0   8/21/2007   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMessageDisplayId, sMessage, sSql, sDisplayField

iMessageDisplayId = request("displayid")
iDisplayOrgId = CLng(request("orgid"))
sMessage = DBsafe(request("message"))
sDisplayField = request("displayfield")

' Clean out the old display
sSql = "Delete FROM egov_organizations_to_displays WHERE orgid = " & iDisplayOrgId & " AND displayid = " & iMessageDisplayId
RunSQL sSql 

' New Refund Policy
sSql = "INSERT INTO egov_organizations_to_displays ( orgid, displayid, " & sDisplayField & " ) VALUES ( "
sSql = sSql & iDisplayOrgId & ", " & iMessageDisplayId & ", '" & sMessage & "' )"
RunSQL sSql 

response.redirect "org_display_edit.asp?orgid=" & iDisplayOrgId & "&displayid=" & iMessageDisplayId


'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' string DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( ByVal strDB )
	Dim sNewString

	If Not VarType( strDB ) = vbString Then DBsafe = strDB : Exit Function

	sNewString = Replace( strDB, "'", "''" )
	'sNewString = Replace( sNewString, "<", "&lt;" )

	DBsafe = sNewString

End Function


'-------------------------------------------------------------------------------------------------
' void RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( ByVal sSql )
	Dim oCmd

	response.write "<p>" & sSql & "</p><br /><br />"
'	response.flush

	Set oCmd = Server.CreateObject("ADODB.Command")
	oCmd.ActiveConnection = Application("DSN")
	oCmd.CommandText = sSql
	oCmd.Execute
	Set oCmd = Nothing

End Sub 


%>