<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: uncomplete.asp
' AUTHOR: Steve Loar
' CREATED: 07/29/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This changes a permit from completed back to the prior status. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   07/29/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iIsOnHold, sInternalComment, sExternalComment, sSql, sActivity, iPermitStatusId, iPriorStatusId
Dim iCurrentStatusID

iPermitId = CLng(request("permitid"))

iCurrentStatusId = GetPermitStatusId( iPermitId )	' in permitcommonfunctions.asp
'iPriorStatusId = GetPriorPermitStatusId( iCurrentStatusId )	' in permitcommonfunctions.asp
iPriorStatusId = GetPermitStatusIdByStatusType( "isissuedback" )

sActivity = "'Permit status changed from " & GetPermitStatusByStatusId( iCurrentStatusId ) & " to " & GetPermitStatusByStatusId( iPriorStatusId ) & "'"

If request("internalcomment") <> "" Then
	sInternalComment = "'" & dbsafe(Trim(request("internalcomment"))) & "'"
Else
	sInternalComment = "NULL"
End If 

If request("externalcomment") <> "" Then 
	sExternalComment = "'" & dbsafe(Trim(request("externalcomment"))) & "'"
Else
	sExternalComment = "NULL"
End If 

sSql = "UPDATE egov_permits SET permitstatusid = " & iPriorStatusId & ", completeddate = NULL WHERE permitid = " & iPermitId
RunSQL sSql

' Push out the expiration date
PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp

'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
MakeAPermitLogEntry iPermitId, "'Permit Notes Added'", sActivity, sInternalComment, sExternalComment, iPriorStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"   ' in permitcommonfunctions.asp

' Set the last activity time to now
SetLastActivityDate( iPermitId )

response.write "UPDATED"


%>