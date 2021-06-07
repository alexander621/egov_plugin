<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitreviewupdate.asp
' AUTHOR: Steve Loar
' CREATED: 06/30/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the permit reviews
'
' MODIFICATION HISTORY
' 1.0   06/30/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitReviewId, sSql, iPermitId, iPermitStatusId, iReviewerId, iReviewStatusId, sInternalNotes
Dim sPublicNotes, iCurrentReviewerId, iCurrentReviewStatusId, sCurrentName, sNewName, sNewStatus
Dim sCurrentStatus, sNote, bReviewChanged, bStatusChanged, bNewNotes, sReturnTo, sSuccessMsg

iPermitReviewId = CLng(request("permitreviewid"))
iPermitId = CLng(request("permitid"))
iPermitStatusId = GetPermitStatusId( iPermitId )		' in permitcommonfunctions.asp
iReviewerId = CLng(request("revieweruserid"))
iReviewStatusId = CLng(request("reviewstatusid"))
sInternalNotes = dbsafe(Trim(request("internalcomment")))
sPublicNotes = dbsafe(Trim(request("externalcomment")))
iCurrentReviewerId = iReviewerId
iCurrentReviewStatusId = iReviewStatusId
bReviewChanged = False 
bStatusChanged = False 
bNewNotes = False 
sNote = "NULL"
sSuccessMsg = "Changes Saved"

' Get the current reviewer and status
GetCurrentReviewerAndStatus iPermitReviewId, iCurrentReviewerId, iCurrentReviewStatusId

If iReviewStatusId = CLng(0) Then
	' the review status will be 0 only when the pick is disabled, which means it should not change
	iReviewStatusId = iCurrentReviewStatusId
End If 

' If the reviewer changed update And make a note
If iCurrentReviewerId <> iReviewerId Then 
	bReviewChanged = True 
	' Get the old reviewer name
	sCurrentName = GetAdminName( iCurrentReviewerId ) ' In common.asp
	If sCurrentName = "" Then 
		sCurrentName = "Unassigned"
	End If 
	' Get the new reviewer name
	
	If CLng(iReviewerId) = CLng(0) Then 
		sNewName = "Unassigned"
		iReviewerId = "NULL"
	Else
		sNewName = GetAdminName( iReviewerId ) ' In common.asp
	End If 

	' Save the reviewer change
	sSql = "UPDATE egov_permitreviews SET revieweruserid = " & iReviewerId
	sSql = sSql & " WHERE permitreviewid = " & iPermitReviewId
	sSql = sSql & " AND permitid = " & iPermitId
	RunSQL sSql

	' email the new reviewer??

	' Create a note
	sNote = "Reviewer changed from " & sCurrentName & " to " & sNewName 
End If 

' If the status changed update and make a note
If iCurrentReviewStatusId <> iReviewStatusId Then 
	bStatusChanged = True 
	' Get the old status
	sCurrentStatus = GetReviewStatusById( iCurrentReviewStatusId )		' in permitcommonfunctions.asp
	' Get the new status
	sNewStatus = GetReviewStatusById( iReviewStatusId )		' in permitcommonfunctions.asp

	' Save the status change
	sSql = "UPDATE egov_permitreviews SET reviewstatusid = " & iReviewStatusId 
	sSql = sSql & ", reviewed = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
	sSql = sSql & " WHERE permitreviewid = " & iPermitReviewId
	sSql = sSql & " AND permitid = " & iPermitId
	RunSQL sSql

	' Create a note
	If sNote <> "NULL" Then
		sNote = sNote & "<br />"
	Else
		sNote = ""
	End If 
	sNote = sNote & "The status changed from " & sCurrentStatus & " to " & sNewStatus

	' Send out emails to all reviewers who get the status change alert for this permit type
	SendReviewStatusChangeAlert iPermitId, iPermitReviewId, sCurrentStatus, sNewStatus

End If 

' Figure out if the user entered notes
If sInternalNotes <> "" Or sPublicNotes <> "" Then 
	bNewNotes = True 
	If sInternalNotes = "" Then
		sInternalNotes = "NULL"
	Else
		sInternalNotes = "'" & sInternalNotes & "'"
	End If 
	If sPublicNotes = "" Then
		sPublicNotes = "NULL"
	Else
		sPublicNotes = "'" & sPublicNotes & "'"
	End If 
Else
	sInternalNotes = "NULL"
	sPublicNotes = "NULL"
End If 

If bNewNotes Or bStatusChanged Or bReviewChanged Then
	If sNote <> "NULL" Then 
		sNote = "'" & sNote & "'"
	End If 
	'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
	MakeAPermitLogEntry iPermitId, "'Permit Review Change'", sNote, sInternalNotes, sPublicNotes, iPermitStatusId, 0, 1, 0, iPermitReviewId, "NULL", iReviewStatusId, "NULL"   ' in permitcommonfunctions.asp

	' Push out the expiration date
	PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp
End If 

' if the permit has not been approved yet 
If Not PermitHasBeenApproved( iPermitId ) Then		' in permitcommonfunctions.asp
	' See if any are not in approved status
	If AllReviewsHaveBeenApproved( iPermitId ) Then 
		' Get the current status
		sCurrentStatus = GetPermitStatusByStatusId( iPermitStatusId )		' in permitcommonfunctions.asp
		
		' Get the approved status
		iNewStatusID = GetPermitStatusIdByStatusType( "isapproved" )		' in permitcommonfunctions.asp
		sNewStatus = GetPermitStatusByStatusId( iNewStatusID )				' in permitcommonfunctions.asp
		
		' Send out an email to all reviewers who get the all approved alert
		SendPermitApprovedAlert iPermitId		' in permitcommonfunctions.asp
	End If 
End If 

sReturnTo = request("reviewpage") & ".asp?permitreviewid=" & iPermitReviewId & "&success=" & sSuccessMsg 
If request("activetab") <> "" Then
	sReturnTo = sReturnTo & "&activetab=" & request("activetab")
End If 

' Go back to the review edit page
response.redirect sReturnTo



'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' Sub GetCurrentReviewerAndStatus( iPermitReviewId, iCurrentReviewerId, iCurrentReviewStatusId )
'-------------------------------------------------------------------------------------------------
Sub GetCurrentReviewerAndStatus( ByVal iPermitReviewId, ByRef iCurrentReviewerId, ByRef iCurrentReviewStatusId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(revieweruserid,0) AS revieweruserid, ISNULL(reviewstatusid,0) AS reviewstatusid "
	sSql = sSql & " FROM egov_permitreviews WHERE permitreviewid = " & iPermitReviewId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iCurrentReviewerId = CLng(oRs("revieweruserid"))
		iCurrentReviewStatusId = CLng(oRs("reviewstatusid"))
	Else
		iCurrentReviewerId = CLng(0)
		iCurrentReviewStatusId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


'-------------------------------------------------------------------------------------------------
' Function AllReviewsHaveBeenApproved( iPermitId )
'-------------------------------------------------------------------------------------------------
Function AllReviewsHaveBeenApproved( iPermitId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitreviewid) AS hits "
	sSql = sSql & " FROM egov_permitreviews R, egov_reviewstatuses S "
	sSql = sSql & " WHERE R.reviewstatusid = S.reviewstatusid AND S.isapproved = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then
			AllReviewsHaveBeenApproved = False  ' Still some that are not approved
		Else
			AllReviewsHaveBeenApproved = True  ' All are approved
		End If 
	Else
		AllReviewsHaveBeenApproved = True   ' No reviews - something is wrong
	End If 

	oRs.Close
	Set oRs = Nothing 
End Function 




%>
