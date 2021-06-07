<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: header_update.asp
' AUTHOR: Steve Loar
' CREATED: 4/11/2007
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates or updates the Receipt Headers
'
' MODIFICATION HISTORY
' 1.0   4/117/07   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iHeaderDisplayId, iFooterDisplayId, sHeader, sFooter, sSql, sRefundFooter, iRefundFooterId

iHeaderDisplayId = request("headerdisplayid")
iFooterDisplayId = request("footerdisplayid")
iRefundFooterId = request("refundfooterid")
sHeader = dbsafe(request("header"))
sFooter = dbsafe(request("footer"))
sRefundFooter = dbsafe(request("refundfooter"))

' Clean out the old header and footer
sSql = "Delete from egov_organizations_to_displays where orgid = " & Session("orgid") & " and displayid = " & iHeaderDisplayId
RunSQL sSql 
sSql = "Delete from egov_organizations_to_displays where orgid = " & Session("orgid") & " and displayid = " & iFooterDisplayId
RunSQL sSql 
sSql = "Delete from egov_organizations_to_displays where orgid = " & Session("orgid") & " and displayid = " & iRefundFooterId
RunSQL sSql 



' New Header
sSql = "Insert into egov_organizations_to_displays ( orgid, displayid, displaydescription ) values ( "
sSql = sSql & Session("orgid") & ", " & iHeaderDisplayId & ", '" & sHeader & "' )"
RunSQL sSql 

' New Footer
sSql = "Insert into egov_organizations_to_displays ( orgid, displayid, displaydescription ) values ( "
sSql = sSql & Session("orgid") & ", " & iFooterDisplayId & ", '" & sFooter & "' )"
RunSQL sSql 

' New Refund Footer
sSql = "Insert into egov_organizations_to_displays ( orgid, displayid, displaydescription ) values ( "
sSql = sSql & Session("orgid") & ", " & iRefundFooterId & ", '" & sRefundFooter & "' )"
RunSQL sSql 


response.redirect "header_edit.asp"


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