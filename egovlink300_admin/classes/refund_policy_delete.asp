<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: refund_policy_delete.asp
' AUTHOR: Steve Loar
' CREATED: 8/10/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes the refund policy, goes with refund_policy_edit.asp
'
' MODIFICATION HISTORY
' 1.0   8/10/07   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMessageDisplayId, sMessage, sSql

iMessageDisplayId = request("messagedisplayid")

' Clean out the old Refund Policy
sSql = "Delete from egov_organizations_to_displays where orgid = " & Session("orgid") & " and displayid = " & iMessageDisplayId
RunSQL sSql 

response.redirect "refund_policy_edit.asp"


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