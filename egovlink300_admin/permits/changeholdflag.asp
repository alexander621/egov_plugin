<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: changeholdflag.asp
' AUTHOR: Steve Loar
' CREATED: 07/22/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This toggles the is on hold flag for a permit. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   07/22/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, iIsOnHold, sInternalComment, sExternalComment, sSql, sActivity, iPermitStatusId

iPermitId = CLng(request("permitid"))
iIsOnHold = CLng(request("isonhold"))

If iIsOnHold = CLng(0) Then
	sActivity = "'Hold removed from permit'"
Else
	sActivity = "'Hold placed on permit'"
End If 

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

sSql = "UPDATE egov_permits SET isonhold = " & iIsOnHold & ", expirationdate = NULL WHERE permitid = " & iPermitId
RunSQL sSql

If iIsOnHold = CLng(0) Then
	' Push out the expiration date
	PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp
End If 

iPermitStatusId = GetPermitStatusId( iPermitId )	' in permitcommonfunctions.asp

'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
MakeAPermitLogEntry iPermitId, "'Permit Notes Added'", sActivity, sInternalComment, sExternalComment, iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"   ' in permitcommonfunctions.asp

response.write "UPDATED"


%>