<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitalertupdate.asp
' AUTHOR: Steve Loar
' CREATED: 08/18/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This creates and updates the permit alerts. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   08/18/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sType, sSql, sAlertMsg, sUserId, sDate, bMessageSet, iPermitStatusId

iPermitId = CLng(request("permitid"))
bMessageSet = False 
'response.write iPermitId & "<br />"

sType = LCase(request("type"))
'response.write sType & "<br />"

If sType = "clear" Then
	sSql = "UPDATE egov_permits SET alertmsg = NULL, alertsetbyuserid = NULL, alertdate = NULL WHERE permitid = " & iPermitId
Else
	'response.write request("alertmsg") & "<br />"
	sUserId = session("userid")
	sDate = "dbo.GetLocalDate(" & Session("OrgID") & ",getdate())"
	bMessageSet = True 

	If request("alertmsg") = "" Then
		sAlertMsg = "NULL"
	Else
		sAlertMsg = "'" & dbsafe(request("alertmsg")) & "'"
	End If 

	sSql = "UPDATE egov_permits SET alertmsg = " & sAlertMsg & ", alertsetbyuserid = " & sUserId
	sSql = sSql & ", alertdate = " & sDate & " WHERE permitid = " & iPermitId
End If 

'response.write sSql & "<br />"
RunSQL sSql 


' get the permit status id for the log entry
iPermitStatusId = GetPermitStatusId( iPermitId )

If bMessageSet Then
	' Put an entry into the log
	'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
	MakeAPermitLogEntry iPermitId, "'Permit Alert Set'", "'Permit alert set to ''" & dbsafe(request("alertmsg")) & "'''", "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"
Else
	' The cleared the message so log that too
	MakeAPermitLogEntry iPermitId, "'Permit Alert Cleared'", "'Permit alert cleared'", "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"
End If 

' Push out the expiration date
PushOutPermitExpirationDate( iPermitId )

response.write "SUCCESS"

%>