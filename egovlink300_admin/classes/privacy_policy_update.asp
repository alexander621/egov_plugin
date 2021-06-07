<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: privacy_policy_update.asp
' AUTHOR: Steve Loar
' CREATED: 7/12/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates or updates the privacy policy, goes with privacy_policy_edit.asp
'
' MODIFICATION HISTORY
' 1.0   7/12/07   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMessageDisplayId, sMessage, sSql

iMessageDisplayId = request("messagedisplayid")
sMessage = DBsafe(request("privacypolicy"))

' Clean out the old Privacy Policy
sSql = "Delete from egov_organizations_to_displays where orgid = " & Session("orgid") & " and displayid = " & iMessageDisplayId
RunSQL sSql 



' New Refund Policy
sSql = "Insert into egov_organizations_to_displays ( orgid, displayid, displaydescription ) values ( "
sSql = sSql & Session("orgid") & ", " & iMessageDisplayId & ", '" & sMessage & "' )"
RunSQL sSql 

response.redirect "privacy_policy_edit.asp"


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------------
' Function DBsafe( strDB )
'------------------------------------------------------------------------------------------------------------
Function DBsafe( strDB )
	If Not VarType( strDB ) = vbString Then 
		DBsafe = strDB 
	Else 
		DBsafe = Replace( strDB, "'", "''" )
	End If 
End Function


'-------------------------------------------------------------------------------------------------
' Sub RunSQL( sSql )
'-------------------------------------------------------------------------------------------------
Sub RunSQL( sSql )
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