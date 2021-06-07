<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: statuschange.asp
' AUTHOR: Steve Loar
' CREATED: 05/21/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This changes the status to the next status for a permit. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   05/21/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql, oRs, sReturn, sDateField, sCurrentStatus, sNewStatus, iCurrentStatusId
Dim iNextStatusId, sNextStatus

iPermitId = CLng(request("permitid"))

iCurrentStatusId = GetPermitStatusId( iPermitId )     ' in permitcommonfunctions.asp
iReqStatusId = GetPermitStatusIdByStatusName(request.querystring("newstatus"))

if iCurrentStatusId = iReqStatusId then
	response.write "SAME"
	response.end
end if

sSql = "SELECT nextpermitstatusid, permitnumberprefix FROM egov_permits P, egov_permitstatuses S "
sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId 
sSql = sSql & " AND P.orgid = " & session("orgid")

Set oRs = Server.CreateObject("ADODB.Recordset")
oRs.Open sSql, Application("DSN"), 3, 1

If Not oRs.EOF Then
	
	If CLng(oRs("nextpermitstatusid")) > CLng(0) And StatusAllowsManualChange( CLng(oRs("nextpermitstatusid")) ) Then 
		' Get the date field to update
		sDateField = GetNextStatusDateField( oRs("nextpermitstatusid") )

		If NextStatusGeneratesThePermitNumber( oRs("nextpermitstatusid") ) Then
			SetPermitNumber iPermitid, oRs("permitnumberprefix")    ' in permitcommonfunctions.asp
		End If 

		' Get the old status
		sCurrentStatus = GetPermitStatusByStatusId( iCurrentStatusId )		' in permitcommonfunctions.asp
		' Get the new status
		sNewStatus = GetPermitStatusByStatusId( oRs("nextpermitstatusid") )		' in permitcommonfunctions.asp


		' Do the update
		sSql = "UPDATE egov_permits SET permitstatusid = " & oRs("nextpermitstatusid") & ", "
		sSql = sSql & sDateField & " = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
		sSql = sSql & " WHERE permitid = " & iPermitId 

		RunSQL sSql

		' Push out the expiration date
		PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp

		'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
		MakeAPermitLogEntry iPermitId, "'Permit Status Changed'", "'Permit status changed from " & sCurrentStatus & " to " & sNewStatus & "'", "NULL", "NULL", oRs("nextpermitstatusid"), 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

		If ReviewersNeedNotification( oRs("nextpermitstatusid") ) Then
'			response.write "Notifying " & oRs("notifyreviewers") & " "
			NotifyReviewers iPermitId, sNewStatus
		End If 



		' Handle Auto Approve due to no Reviews here. This is for Released to Approved status change only.
		If StatusNeedsReviewsToChange( oRs("nextpermitstatusid") ) Then
			If PermitHasNoReviews( iPermitId ) Then
				' Get the next status
				iNextStatusId = GetNextPermitStatusId( iPermitId )

				' Get the name of that status
				sNextStatus = GetPermitStatusByStatusId( iNextStatusId )		' in permitcommonfunctions.asp

				' Get the date field to update
				sDateField = GetNextStatusDateField( iNextStatusId )

				' update the permit to the next status
				sSql = "UPDATE egov_permits SET permitstatusid = " & iNextStatusId & ", "
				sSql = sSql & sDateField & " = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
				sSql = sSql & " WHERE permitid = " & iPermitId 
				RunSQL sSql

				'Make A Permit Log Entry
				MakeAPermitLogEntry iPermitId, "'Permit Status Changed'", "'Permit status was automatically changed from " & sNewStatus & " to " & sNextStatus & " because there are no associated reviews for this permit.'", "NULL", "NULL", iNextStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

				' Send the permit approved alerts if any are set
				SendPermitApprovedAlert iPermitId 		' in permitcommonfunctions.asp
			End If 
		End If 

		' Handle Auto complete due to no inspections here. This is for Issued to Completed status change only.
		' Get the next status
		iNextStatusId = GetNextPermitStatusId( iPermitId )
		If ( GetPermitStatusIdByStatusType( "iscompletedstatus" ) = iNextStatusId ) And ( iNextStatusId <> CLng(0) ) Then 
			If PermitHasNoInspections( iPermitId ) Then 
				' Get the name of that status
				sNextStatus = GetPermitStatusByStatusId( iNextStatusId )		' in permitcommonfunctions.asp

				' Get the date field to update
				sDateField = GetNextStatusDateField( iNextStatusId )

				' update the permit to the next status (completed)
				sSql = "UPDATE egov_permits SET permitstatusid = " & iNextStatusId & ", "
				sSql = sSql & sDateField & " = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) "
				sSql = sSql & " WHERE permitid = " & iPermitId 
				RunSQL sSql

				'Make A Permit Log Entry
				MakeAPermitLogEntry iPermitId, "'Permit Status Changed'", "'Permit status was automatically changed from " & sNewStatus & " to " & sNextStatus & " because there are no associated inspections for this permit.'", "NULL", "NULL", iNextStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"
			End If 
		End If 

		' Set the last activity date to now
		SetLastActivityDate( iPermitId )

		sReturn = "UPDATED"
	Else 
		sReturn = "SAME"
	End If 
Else
	sReturn = "SAME"
End If 

oRs.Close
Set oRs = Nothing 

response.write sReturn



'--------------------------------------------------------------------------------------------------
' SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' string GetNextStatusDateField( iStatusId )
'-------------------------------------------------------------------------------------------------
Function GetNextStatusDateField( ByVal iStatusId )
	Dim sSql, oRs

	sSql = "SELECT statusdatedisplayed FROM egov_permitstatuses WHERE permitstatusid = " & iStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetNextStatusDateField = oRs("statusdatedisplayed")
	Else
		GetNextStatusDateField = "applieddate"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean NextStatusGeneratesThePermitNumber( iStatusId )
'-------------------------------------------------------------------------------------------------
Function NextStatusGeneratesThePermitNumber( ByVal iStatusId )
	Dim sSql, oRs

	sSql = "SELECT generatespermitno FROM egov_permitstatuses WHERE permitstatusid = " & iStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		NextStatusGeneratesThePermitNumber = oRs("generatespermitno")
	Else
		NextStatusGeneratesThePermitNumber = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean ReviewersNeedNotification( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function ReviewersNeedNotification( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT notifyreviewers FROM egov_permitstatuses WHERE permitstatusid = " & iPermitStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("notifyreviewers") Then 
			ReviewersNeedNotification = True 
		Else
			ReviewersNeedNotification = False 
		End If 
	Else
		ReviewersNeedNotification = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' void NotifyReviewers iPermitId, sNewStatus 
'-------------------------------------------------------------------------------------------------
Sub NotifyReviewers( ByVal iPermitId, ByVal sNewStatus )
	Dim sSql, oRs, sToName, sToEmail, sFromName, sFromEmail, sSubject, sHTMLBody
	Dim sPermitNo, sDesc, sJobSite, iPermitTypeId, sOrgName, sStatus, sLocation
	Dim sLocationType

	' Pull the permit details needed
	sPermitNo = GetPermitNumber( iPermitId )
	sDesc = GetPermitTypeDesc( iPermitId, True ) '	in permitcommonfunctions.asp
	sJobSite = GetPermitJobSite( iPermitId )
	iPermitTypeId = GetPermitTypeId( iPermitId )
	sStatus = GetPermitStatusByPermitId( iPermitId ) '	in permitcommonfunctions.asp
	sLocation = Replace(GetPermitPermitLocation( iPermitId ), Chr(10), Chr(10) & "<br />")
	sLocationType = GetPermitLocationType( iPermitId )

	sSubject = "Permit " & sPermitNo & " " & sNewStatus
	sOrgName = GetOrgName( session("orgid") )
	sFromName = sOrgName & " E-GOV WEBSITE"

	' Build the email body
	sHTMLBody = "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & vbcrlf & vbcrlf & "<p>This permit has been " & sNewStatus & ".</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & vbcrlf & vbcrlf & "<p>Permit #: " & sPermitNo & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & vbcrlf & "Permit Type: " & sDesc & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & vbcrlf & "Permit Status: " & sStatus & "<br />" & vbcrlf
	If sLocationType = "address" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Job Site: " & sJobSite
	End If 
	If sLocationType = "location" Then 
		sHTMLBody = sHTMLBody & vbcrlf & "Location: " & sLocation
	End If 
	sHTMLBody = sHTMLBody & "</p>" & vbcrlf & vbcrlf
	sHTMLBody = sHTMLBody & vbcrlf & vbcrlf & "<p>Click here to view the reviews for this permit.<br />"
	sHTMLBody = sHTMLBody & vbcrlf & "<a href=""" & session("egovclientwebsiteurl") & "/admin/permits/permitreviewerlist.asp?permitid=" & iPermitId & """ title=""click to view"">" & session("egovclientwebsiteurl") & "/admin/permits/permitreviewerlist.asp?permitid=" & iPermitId & "</a></p>"

	' Pull any reviewers that reviewers for this permit
	sSql = "SELECT DISTINCT U.firstname, U.lastname, ISNULL(U.email,'') AS email "
	sSql = sSql & " FROM users U, egov_permitreviews R "
	sSql = sSql & " WHERE U.userid = R.revieweruserid AND u.isdeleted = 0 AND R.notifyonrelease = 1 AND R.permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		If oRs("email") <> "" Then 
			sToName = oRs("firstname") & " " & oRs("lastname") ' -- Original code
			'SendEmailPermits sToName, oRs("email"), sFromName, "webmaster@eclink.com", sSubject, sHTMLBody		' in permitcommonfunctions.asp

			'sendEmail "", oRs("email") & "(" & sToName & ")", "dboyer@eclink.com (David Boyer)", sSubject, sHTMLBody, "", ""   ' -- Original code
			sendEmail "", sToName & " <" & oRs("email") & ">", "", sSubject, sHTMLBody, "", "" 
		End If 
		oRs.MoveNext 
	Loop
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'-------------------------------------------------------------------------------------------------
' boolean StatusNeedsReviewsToChange( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function StatusNeedsReviewsToChange( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT needsreviewstochange FROM egov_permitstatuses WHERE permitstatusid = " & iPermitStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("needsreviewstochange") Then 
			StatusNeedsReviewsToChange = True 
		Else
			StatusNeedsReviewsToChange = False 
		End If 
	Else
		StatusNeedsReviewsToChange = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean StatusAllowsManualChange( iPermitStatusId )
'-------------------------------------------------------------------------------------------------
Function StatusAllowsManualChange( ByVal iPermitStatusId )
	Dim sSql, oRs

	sSql = "SELECT canmanuallyset FROM egov_permitstatuses WHERE permitstatusid = " & iPermitStatusId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("canmanuallyset") Then 
			StatusAllowsManualChange = True 
		Else
			StatusAllowsManualChange = False 
		End If 
	Else
		StatusAllowsManualChange = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' boolean PermitHasNoReviews( iPermitId )
'-------------------------------------------------------------------------------------------------
Function PermitHasNoReviews( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitreviewid) AS hits FROM egov_permitreviews WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasNoReviews = False  
		Else
			PermitHasNoReviews = True  
		End If 
	Else
		PermitHasNoReviews = True  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' byval  PermitHasNoInspections( iPermitId )
'-------------------------------------------------------------------------------------------------
Function PermitHasNoInspections( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(permitinspectionid) AS hits FROM egov_permitinspections WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If CLng(oRs("hits")) > CLng(0) Then 
			PermitHasNoInspections = False  
		Else
			PermitHasNoInspections = True  
		End If 
	Else
		PermitHasNoInspections = True  
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' integer GetNextPermitStatusId( iPermitId )
'-------------------------------------------------------------------------------------------------
Function GetNextPermitStatusId( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT nextpermitstatusid FROM egov_permits P, egov_permitstatuses S "
	sSql = sSql & " WHERE P.permitstatusid = S.permitstatusid AND P.permitid = " & iPermitId 
	sSql = sSql & " AND P.orgid = " & session("orgid")

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		GetNextPermitStatusId = CLng(oRs("nextpermitstatusid"))
	Else
		GetNextPermitStatusId = CLng(0)
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function



%>
