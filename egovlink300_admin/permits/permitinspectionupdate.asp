<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitinspectionupdate.asp
' AUTHOR: Steve Loar
' CREATED: 07/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates the permit inspections
'
' MODIFICATION HISTORY
' 1.0   07/11/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitInspectionId, sSql, iPermitId, iPermitStatusId, iPermitTypeId, iInspectorUserId, iInspectionStatusId
Dim sRequestReceived, sRequestedDate, sRequestedTime, sRequestedAmPm, sScheduledDate, sScheduledTime
Dim sScheduledAmPm, sInspectedDate, sInspectedTime, sInspectedAmPm, sContactPhone, sContact, sSchedulingNotes
Dim iPermitInspectionTypeId, bIsFinal, iCurrentInspectorUserId, iCurrentInspectionStatusId, sCurrentRequestedDate
Dim sCurrentRequestedTime, sCurrentRequestedAmPm, sCurrentScheduledDate, sCurrentScheduledTime, sCurrentScheduledAmPm
Dim sCurrentInspectedDate, sCurrentInspectedTime, sCurrentInspectedAmPm, sCurrentContactPhone, sCurrentContact
Dim sCurrentSchedulingNotes, sInternalNotes, sPublicNotes, sCurrentName, sNewName, bInspectorChanged, bStatusChanged
Dim bNewNotes, sNote, sCurrentStatus, sNewStatus, bChanges, iNewPermitInspectionId, iPassedStatusId, sSuccessMsg
Dim sInspectionDescription, sMakeics

' Grab the data from the form post
iPermitId = CLng(request("permitid"))
iPermitInspectionId = CLng(request("permitinspectionid"))
iPermitStatusId = GetPermitStatusId( iPermitId )		' in permitcommonfunctions.asp
iInspectorUserId = CLng(request("inspectoruserid"))
iInspectionStatusId = CLng(request("inspectionstatusid")) ' The new status that they picked for this inspection
sRequestedDate = request("requesteddate")
sRequestedTime = request("requestedtime") 
sRequestedAmPm = request("requestedampm")
sScheduledDate = request("scheduleddate")
sScheduledTime = request("scheduledtime")
sScheduledAmPm = request("scheduledampm")
sInspectedDate = request("inspecteddate")
sInspectedTime = request("inspectedtime")
sInspectedAmPm = request("inspectedampm")
sContact = dbsafe(request("contact"))
sContactPhone = dbsafe(request("contactphone"))
sSchedulingNotes = dbsafe(request("schedulingnotes"))
sInternalNotes = dbsafe(Trim(request("internalcomment")))
sPublicNotes = dbsafe(Trim(request("externalcomment")))
sSuccessMsg = "Changes Saved"
sInspectionDescription = ""
smakeics = request("makeics")

' Set some current values to form post before fetch of actual current values
iCurrentInspectorUserId = iInspectorUserId
iCurrentInspectionStatusId = iInspectionStatusId
bInspectorChanged = False 
bStatusChanged = False 
bNewNotes = False 
sNote = "NULL"
bChanges = False 

' Get current inspector, status, notes, requested date/time, scheduled date/time, inspected date/time
GetCurrentInspectionDetails iPermitInspectionId

response.write "iCurrentInspectorUserId = " & iCurrentInspectorUserId & "<br /><br />"

' If the inspector has changed
If CLng(iCurrentInspectorUserId) <> CLng(iInspectorUserId) Then 
	bInspectorChanged = True 
	' Get the old inspector name
	sCurrentName = GetAdminName( iCurrentInspectorUserId ) ' In common.asp
	If sCurrentName = "" Then 
		sCurrentName = "Unassigned"
	End If 
	' Get the new inspector name
	sNewName = GetAdminName( iInspectorUserId ) ' In common.asp

	' Save the inspector change
	sSql = "UPDATE egov_permitinspections SET inspectoruserid = " & iInspectorUserId
	sSql = sSql & " WHERE permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " AND permitid = " & iPermitId
	response.write sSql & "<br /><br />"
	RunSQL sSql

	' Create a note
	sNote = "Inspector changed from " & sCurrentName & " to " & sNewName 
End If 


' If the status has changed save that change
If iCurrentInspectionStatusId <> iInspectionStatusId Then 
	bStatusChanged = True 
	' Get the old status
	sCurrentStatus = GetInspectionStatusById( iCurrentInspectionStatusId )		' in permitcommonfunctions.asp
	' Get the new status
	sNewStatus = GetInspectionStatusById( iInspectionStatusId )		' in permitcommonfunctions.asp

	' Save the status change
	sSql = "UPDATE egov_permitinspections SET inspectionstatusid = " & iInspectionStatusId 
	sSql = sSql & " WHERE permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " AND permitid = " & iPermitId
	response.write sSql & "<br /><br />"
	RunSQL sSql

	' Create a note
	If sNote <> "NULL" Then
		sNote = sNote & "<br />"
	Else
		sNote = ""
	End If 
	sNote = sNote & "The status changed from " & sCurrentStatus & " to " & sNewStatus
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

If bNewNotes Or bStatusChanged Or bInspectorChanged Then
	bChanges = True 
	If sNote <> "NULL" Then 
		sNote = "'" & sNote & "'"
	End If 
	'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
	MakeAPermitLogEntry iPermitId, "'Permit Inspection Change'", sNote, sInternalNotes, sPublicNotes, iPermitStatusId, 1, 0, 0, "NULL", iPermitInspectionId, "NULL", iInspectionStatusId   ' in permitcommonfunctions.asp
End If 

response.write "Through Notes<br /><br />"

' check for new schedule dates
If sRequestReceived = "" And (sRequestedDate <> "" Or sScheduledDate <> "" Or sInspectedDate <> "") Then 
	bChanges = True 
	sSql = "UPDATE egov_permitinspections SET requestreceiveddate = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()) " 
	sSql = sSql & " WHERE permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " AND permitid = " & iPermitId
	response.write sSql & "<br /><br />"
	RunSQL sSql
End If 

' check other fields for any changes and update the data
If sRequestedDate <> sCurrentRequestedDate Or sRequestedTime <> sCurrentRequestedTime Or sRequestedAmPm <> sCurrentRequestedAmPm _
	Or sScheduledDate <> sCurrentScheduledDate Or sScheduledTime <> sCurrentScheduledTime Or sScheduledAmPm <> sCurrentScheduledAmPm _
	Or sInspectedDate <> sCurrentInspectedDate Or sInspectedTime <> sCurrentInspectedTime Or sInspectedAmPm <> sCurrentInspectedAmPm Or _
	sContact <> sCurrentContact Or sContactPhone <> sCurrentContactPhone Or sSchedulingNotes <> sCurrentSchedulingNotes Then 

	bChanges = True

	' Prep the data for the update
	If sRequestedDate <> "" Then
		sRequestedDate = "'" & sRequestedDate & "'"
		If sRequestedTime <> "" Then
			sRequestedTime = "'" & sRequestedTime & "'"
			sRequestedAmPm = "'" & sRequestedAmPm & "'"
		Else
			sRequestedTime = "NULL"
			sRequestedAmPm = "NULL"
		End If 
	Else
		sRequestedDate = "NULL"
		sRequestedTime = "NULL"
		sRequestedAmPm = "NULL"
	End If 

	If sScheduledDate <> "" Then
		sScheduledDate = "'" & sScheduledDate & "'"
		If sScheduledTime <> "" Then
			sScheduledTime = "'" & sScheduledTime & "'"
			sScheduledAmPm = "'" & sScheduledAmPm & "'"
		Else
			sScheduledTime = "NULL"
			sScheduledAmPm = "NULL"
		End If 
	Else
		sScheduledDate = "NULL"
		sScheduledTime = "NULL"
		sScheduledAmPm = "NULL"
	End If 

	If sInspectedDate <> "" Then
		sInspectedDate = "'" & sInspectedDate & "'"
		If sInspectedTime <> "" Then
			sInspectedTime = "'" & sInspectedTime & "'"
			sInspectedAmPm = "'" & sInspectedAmPm & "'"
		Else
			sInspectedTime = "NULL"
			sInspectedAmPm = "NULL"
		End If 
	Else
		sInspectedDate = "NULL"
		sInspectedTime = "NULL"
		sInspectedAmPm = "NULL"
	End If 

	If sContact <> "" Then
		sContact = "'" & sContact & "'"
	Else
		sContact = "NULL"
	End If 

	If sContactPhone <> "" Then
		sContactPhone = "'" & sContactPhone & "'"
	Else
		sContactPhone = "NULL"
	End If 

	If sSchedulingNotes <> "" Then
		sSchedulingNotes = "'" & sSchedulingNotes & "'"
	Else
		sSchedulingNotes = "NULL"
	End If 

	' Save changes
	sSql = "UPDATE egov_permitinspections SET "
	sSql = sSql & " requesteddate = " & sRequestedDate & ","
	sSql = sSql & " requestedtime = " & sRequestedTime & ","
	sSql = sSql & " requestedampm = " & sRequestedAmPm & ","
	sSql = sSql & " scheduleddate = " & sScheduledDate & ","
	sSql = sSql & " scheduledtime = " & sScheduledTime & ","
	sSql = sSql & " scheduledampm = " & sScheduledAmPm & ","
	sSql = sSql & " inspecteddate = " & sInspectedDate & ","
	sSql = sSql & " inspectedtime = " & sInspectedTime & ","
	sSql = sSql & " inspectedampm = " & sInspectedAmPm & ","
	sSql = sSql & " contact = " & sContact & ","
	sSql = sSql & " contactphone = " & sContactPhone & ","
	sSql = sSql & " schedulingnotes = " & sSchedulingNotes
	sSql = sSql & " WHERE permitinspectionid = " & iPermitInspectionId
	sSql = sSql & " AND permitid = " & iPermitId
	response.write sSql & "<br /><br />"
	RunSQL sSql

	If sInspectionDescription <> sCurrentScheduledDate And sCurrentScheduledDate = "" Then
		' If this is being scheduled then see if the applicant is to be alerted
		If AlertApplciantWhenScheduled( iPermitId ) Then
			SendInspectionScheduledAlert iPermitId, sInspectionDescription, "applicantuserid", sScheduledDate, sScheduledTime, sScheduledAmPm
		End If 
	End If 
End If 

If bChanges Then 
	' Push out the expiration date
	PushOutPermitExpirationDate iPermitId   ' in permitcommonfunctions.asp

	' Set the last activity date for the permit
	SetLastActivityDate( iPermitId )
End If 

'response.write "Past push out <br /><br />"

If bStatusChanged Then 
	' Does the new status require a reschedule 
	If RescheduleInspectionForStatus( iInspectionStatusId ) Then		' in permitcommonfunctions.asp
		' add this inspection back to the list with the resched flag set
		iNewPermitInspectionId = ReschedulePermitInspection( iPermitInspectionId )		' in permitcommonfunctions.asp

		' Get the new status of this new inspection
		iInitialStatusid = GetInspectionStatusId( "isinitialstatus" )	' in permitcommonfunctions.asp

		' Add a log entry for this new inspection
		'MakeAPermitLogEntry( iPermitid, sActivity, sActivityComment, sInternalComment, sExternalComment, iPermitStatusId, iIsInspectionEntry, iIsReviewEntry, iIsActivityEntry, iPermitReviewId, iPermitInspectionId, iReviewStatusId, iInspectionStatusId )
		MakeAPermitLogEntry iPermitId, "'Permit Reinspection Added'", "'Reinspection added to permit'", "NULL", "NULL", iPermitStatusId, 1, 0, 0, "NULL", iNewPermitInspectionId, "NULL", iInitialStatusid   ' in permitcommonfunctions.asp

		'Copy Inspection Report Data
		sSQL = "SELECT inspectiontype,permitinspectionreportid FROM egov_permitinspectionreports WHERE permitinspectionid = " & iPermitInspectionId
		set oIR = Server.CreateObject("ADODB.RecordSet")
		oIR.Open sSQL, Application("DSN"), 3, 1
		if not oIR.EOF then
			sSQL = "INSERT INTO egov_permitinspectionreports (permitid,permitinspectionid,inspectiontype) " _
				& " VALUES(" _
				& "'" & iPermitId & "'," _
				& "'" & iNewPermitInspectionId & "'," _
				& "'" & oIR("inspectiontype") & "')"
			'response.write sSQL & "<br />"
			intPermitInspectionReportID = RunInsertStatement(sSQL)

			sSQL = "SELECT inspectiontype FROM egov_permitinspectionreporttypes WHERE permitinspectionreportid = " & oIR("permitinspectionreportid")
			set oIT = Server.CreateObject("ADODB.RecordSet")
			oIT.Open sSQL, Application("DSN"), 3, 1
			Do While NOT oIT.EOF
				sSQL = "INSERT INTO egov_permitinspectionreporttypes (permitinspectionreportid,inspectiontype) VALUES('" & intPermitInspectionReportID & "','" & oIT("inspectiontype") & "')"
				'response.write sSQL & "<br />"
				RunSQLStatement(sSQL)
				
				oIT.MoveNext
			loop
			oIT.Close
			Set oIT = Nothing

			MakeAPermitLogEntry iPermitId, "'Permit Inspection Report Created'", "'Inspection Report added to permit'", "NULL", "NULL", iPermitStatusId, 1, 0, 0, "NULL", iNewPermitInspectionId, "NULL", iInitialStatusid   ' in permitcommonfunctions.asp

		end if
		oIR.Close
		Set oIR = Nothing

	Else 
		' Get the statusid for when an inspection has passed
		iPassedStatusId = GetInspectionStatusId( "ispassed" )

		'if request.cookies("user")("userid") then response.end
		If bIsFinal Then 
			' the status of the inspection has changed, and it is the final inspection
			' see if they passed their final inspection
			If CLng(iPassedStatusId) = CLng(iInspectionStatusId) Then 
				' Get the current permit status
				sCurrentStatus = GetPermitStatusByStatusId( iPermitStatusId )		' in permitcommonfunctions.asp
				
				' Get the completed permit status
				iNewStatusID = GetPermitStatusIdByStatusType( "iscompletedstatus" )		' in permitcommonfunctions.asp
				' If they have a completed status then finish processing it.
				If CLng(iNewStatusID) > CLng(0) Then 
					' if the permit status is not already completed, change it
					If CLng(iNewStatusID) <> CLng(iPermitStatusId) Then 
						sNewStatus = GetPermitStatusByStatusId( iNewStatusID )		' in permitcommonfunctions.asp

						' update the permit status to completed, set the completed date, and null out the expiration date
						'sSql = "UPDATE egov_permits SET completeddate = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), permitstatusid = " & iNewStatusID & ", isexpired = 0, expirationdate = NULL WHERE permitid = " & iPermitId
						'RunSQL sSql

						' make a log of the status change on the permit
						'MakeAPermitLogEntry iPermitId, "'Permit Status Changed'", "'Permit status changed from " & sCurrentStatus & " to " & sNewStatus & "'", "NULL", "NULL", iNewStatusID, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

						' send alerts if needed.
						SendPermitPassedFinalInspectionAlert iPermitId		' in permitcommonfunctions.asp
					End If 
				End If 
			End If 
		Else
			' This is not the final inspection but if this inspection passed and all inspections have passed then complete the permit
			If CLng(iPassedStatusId) = CLng(iInspectionStatusId) Then
				If PermitHasNoPendingInspections( iPermitid, iPassedStatusId ) Then
					' Get the current permit status name
					sCurrentStatus = GetPermitStatusByStatusId( iPermitStatusId )		' in permitcommonfunctions.asp
					
					' Get the completed permit status id
					iNewStatusID = GetPermitStatusIdByStatusType( "iscompletedstatus" )		' in permitcommonfunctions.asp

					' if the permit status is not already completed, change it
					If CLng(iNewStatusID) <> CLng(iPermitStatusId) Then 

						' Get the name of the completed status
						sNewStatus = GetPermitStatusByStatusId( iNewStatusID )		' in permitcommonfunctions.asp

						' update the permit status to completed, set the completed date, and null out the expiration date
						'sSql = "UPDATE egov_permits SET completeddate = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), permitstatusid = " & iNewStatusID & ", expirationdate = NULL WHERE permitid = " & iPermitId
						'RunSQL sSql

						' make a log of the status change on the permit
						'MakeAPermitLogEntry iPermitId, "'Permit Status Changed'", "'Permit status automatically changed from " & sCurrentStatus & " to " & sNewStatus & " because all inspections for this permit have passed.'", "NULL", "NULL", iNewStatusID, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

						' send alerts that the permit has passed all inspections
						SendPermitPassedFinalInspectionAlert iPermitId		' in permitcommonfunctions.asp
					End If 
				End If 
			End If 
		End If 
	End If 
Else
	' This code allows a permit that has been pushed back to issued status, to be completed again.
	' see if this inspection is the final and is passed and permit is not completed
	' If so then complete the permit
	If bIsFinal Then 
		iPassedStatusId = GetInspectionStatusId( "ispassed" )

		' see if they passed their final inspection
		If CLng(iPassedStatusId) = CLng(iInspectionStatusId) Then 
			' Get the current permit status
			sCurrentStatus = GetPermitStatusByStatusId( iPermitStatusId )		' in permitcommonfunctions.asp
			
			' Get the completed permit status
			iNewStatusID = GetPermitStatusIdByStatusType( "iscompletedstatus" )		' in permitcommonfunctions.asp
			' If they have a completed status then finish processing it.
			If CLng(iNewStatusID) > CLng(0) Then 
				' if the permit status is not already completed, change it
				If CLng(iNewStatusID) <> CLng(iPermitStatusId) Then 
					sNewStatus = GetPermitStatusByStatusId( iNewStatusID )		' in permitcommonfunctions.asp

					' update the permit status to completed, set the completed date, and null out the expiration date
					'sSql = "UPDATE egov_permits SET completeddate = dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), permitstatusid = " & iNewStatusID & ", expirationdate = NULL WHERE permitid = " & iPermitId
					'RunSQL sSql
					'response.write sSql & "<br />"
					'response.end

					' make a log of the status change on the permit
					'MakeAPermitLogEntry iPermitId, "'Permit Status Changed'", "'Permit status changed from " & sCurrentStatus & " to " & sNewStatus & "'", "NULL", "NULL", iNewStatusID, 0, 0, 1, "NULL", "NULL", "NULL", "NULL"

					' send alerts if needed.
					SendPermitPassedFinalInspectionAlert iPermitId		' in permitcommonfunctions.asp
				End If 
			End If 
		End If 
	End If 
End If 


if session("orgid") = "139" or session("orgid") = "181"  then 
	'Process Permit Inspection Report Fields
	intPermitInspectionReportID = clng(dbsafe(request.form("permitinspectionreportid")))
	'LOOP FOR "CHECKBOXES"
	strApproved = "0"
	if request.form("approved") = "on" then strApproved = "1"
	strDisapproved = "0"
	if request.form("disapproved") = "on" then strDisapproved = "1"
	strApprovedWCorr = "0"
	if request.form("approvedwcorr") = "on" then strApprovedWCorr = "1"
	strCOC = "0"
	if request.form("coc") = "on" then strCOC = "1"

	
	if intPermitInspectionReportID = "0" and (strApproved <> "0" or strDisapproved <> "0" or strApprovedWCorr <> "0" or strCOC <> "0" or request.form("inspectiontype") <> "" or request.form("remarks") <> "" or request.form("permitinspectorid") <> "0") then
		sSQL = "INSERT INTO egov_permitinspectionreports (permitid,permitinspectionid,inspectiontype,approved,disapproved,approvedwcorr,coc,remarks, permitinspectorid) " _
			& " VALUES(" _
			& "'" & iPermitId & "'," _
			& "'" & iPermitInspectionId & "'," _
			& "'" & dbsafe(request.form("inspectiontype")) & "'," _
			& strApproved & "," _
			& strDisapproved & "," _
			& strApprovedWCorr & "," _
			& strCOC & "," _
			& "'" & dbsafe(request.form("remarks")) & "'," _
			& "'" & dbsafe(request.form("permitinspectorid")) & "')"
		'response.write sSQL & "<br />"
		intPermitInspectionReportID = RunInsertStatement(sSQL)

		'Add Permit Note that Permit Inspection Report was created
		MakeAPermitLogEntry iPermitId, "'Permit Inspection Report Created'", "'Inspection Report added to permit'", "NULL", "NULL", iPermitStatusId, 1, 0, 0, "NULL", iPermitInspectionId, "NULL", iInspectionStatusId   ' in permitcommonfunctions.asp

		blnNewRecord = true
	elseif intPermitInspectionReportID <> "0" then
		'Update if other
		sSQL = "UPDATE egov_permitinspectionreports SET " _
			& " inspectiontype = '" & dbsafe(request.form("inspectiontype")) & "'," _
			& " remarks = '" & dbsafe(request.form("remarks")) & "'," _
			& " permitinspectorid = '" & dbsafe(request.form("permitinspectorid")) & "'," _
			& " approved = " & strApproved & "," _
			& " disapproved = " & strDisapproved & "," _
			& " approvedwcorr = " & strApprovedWCorr & "," _
			& " coc = " & strCOC _
			& " WHERE permitid = '" & iPermitId & "' AND permitinspectionreportid = '" & intPermitInspectionReportID & "'"
		'response.write sSQL & "<br />"
		'response.flush
		'response.end
		RunSQLStatement(sSQL)
	
	
	
		'delete any inspectionreporttypes
		sSQL = "DELETE FROM egov_permitinspectionreporttypes WHERE permitinspectionreportid = '" & intPermitInspectionReportID & "'"
		'response.write sSQL & "<br />"
		RunSQLStatement(sSQL)
	end if
	
	if intPermitInspectionReportID <> "0" then
		'loop through fields and insert new inspectionreporttypes
		for each item in request.form
			if instr(item,"inspectionreporttype") = 1 then
				if request.form(item) <> "0" then
					sSQL = "INSERT INTO egov_permitinspectionreporttypes (permitinspectionreportid,inspectiontype) VALUES('" & intPermitInspectionReportID & "','" & request.form(item) & "')"
					'response.write sSQL & "<br />"
					RunSQLStatement(sSQL)
				end if
			end if
		next
	end if
end if







'response.write "At End!"
'response.end

sReturnTo = request("inspectionpage") & ".asp?permitinspectionid=" & iPermitInspectionId & "&success=" & sSuccessMsg 
If request("activetab") <> "" Then
	sReturnTo = sReturnTo & "&activetab=" & request("activetab")
End If 
if request("makeics") = "yes" then
	sReturnTo = sReturnTo & "&makeics=yes"
end if

response.redirect sReturnTo




'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'-------------------------------------------------------------------------------------------------
' GetCurrentInspectionDetails iPermitInspectionId 
'-------------------------------------------------------------------------------------------------
Sub GetCurrentInspectionDetails( ByVal iPermitInspectionId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(inspectoruserid,0) AS inspectoruserid, ISNULL(inspectionstatusid,0) AS inspectionstatusid, isfinal, "
	sSql = sSql & " requestreceiveddate, requesteddate, ISNULL(requestedtime,'') AS requestedtime, ISNULL(requestedampm,'') AS requestedampm, "
	sSql = sSql & " scheduleddate, ISNULL(scheduledtime,'') AS scheduledtime, ISNULL(scheduledampm,'') AS scheduledampm, "
	sSql = sSql & " inspecteddate, ISNULL(inspectedtime,'') AS inspectedtime, ISNULL(inspectedampm,'') AS inspectedampm, "
	sSql = sSql & " contact, contactphone, schedulingnotes, permittypeid, permitinspectiontypeid, ISNULL(inspectiondescription,'') AS inspectiondescription "
	sSql = sSql & " FROM egov_permitinspections WHERE permitinspectionid = " & iPermitInspectionId

	response.write sSql & "<br /><br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		iCurrentInspectorUserId = CLng(oRs("inspectoruserid"))
		iCurrentInspectionStatusId = CLng(oRs("inspectionstatusid"))
		If IsNull(oRs("requestreceiveddate")) Then
			sRequestReceived = ""
		Else
			sRequestReceived = oRs("requestreceiveddate")
		End If 
		If IsNull(oRs("requesteddate")) Then
			sCurrentRequestedDate = ""
		Else 
			sCurrentRequestedDate = oRs("requesteddate")
		End If 
		sCurrentRequestedTime = oRs("requestedtime")
		sCurrentRequestedAmPm = oRs("requestedampm")
		If IsNull(oRs("scheduleddate")) Then 
			sCurrentScheduledDate = ""
		Else 
			sCurrentScheduledDate = oRs("scheduleddate")
		End If 
		sCurrentScheduledTime = oRs("scheduledtime")
		sCurrentScheduledAmPm = oRs("scheduledampm")
		If IsNull(oRs("inspecteddate")) Then
			sCurrentInspectedDate = ""
		Else 
			sCurrentInspectedDate = oRs("inspecteddate")
		End If 
		sCurrentInspectedtime = oRs("inspectedtime")
		sCurrentInspectedAmPm = oRs("inspectedampm")
		sCurrentContactPhone = oRs("contactphone")
		sCurrentContact = oRs("contact")
		sCurrentSchedulingNotes = oRs("schedulingnotes")
		iPermitTypeId = oRs("permittypeid")
		iPermitInspectionTypeId = oRs("permitinspectiontypeid")
		bIsFinal = oRs("isfinal")
		sInspectionDescription = oRs("inspectiondescription")
	End If 

	oRs.Close
	Set oRs = Nothing 
End Sub 


' boolean = AlertApplciantWhenScheduled( iPermitId )
'-------------------------------------------------------------------------------------------------
Function AlertApplciantWhenScheduled( ByVal iPermitId )
	Dim oRs, sSql

	sSql = "SELECT alertapplicantofinspections FROM egov_permits WHERE permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("alertapplicantofinspections") Then
			AlertApplciantWhenScheduled = True 
		Else
			AlertApplciantWhenScheduled = False 
		End If 
	Else
		AlertApplciantWhenScheduled = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 


'-------------------------------------------------------------------------------------------------
' SendInspectionScheduledAlert iPermitId, sInspectionDescription, sToUserId, sScheduledDate, sScheduledTime, sScheduledAmPm
'-------------------------------------------------------------------------------------------------
Sub SendInspectionScheduledAlert( ByVal iPermitId, ByVal sInspectionDescription, ByVal sToUserIdType, ByVal sScheduledDate, ByVal sScheduledTime, ByVal sScheduledAmPm )
	Dim sSql, oRs, sToName, sToEmail, sSubject, sHTMLBody, sFromName, sPermitNo, sDesc
	Dim sJobSite, sOrgName

	' Pull the permit details needed
	sPermitNo = GetPermitNumber( iPermitId )
	sDesc = GetPermitTypeDesc( iPermitId, False )
	sJobSite = GetPermitJobSite( iPermitId )
	'iPermitTypeId = GetPermitTypeId( iPermitId )

	sSubject = "Permit " & sPermitNo & " - An Inspection Has Been Scheduled"
	sOrgName = GetOrgName( session("orgid") )
	sFromName = sOrgName & " E-GOV WEBSITE"

	' Build the email body
	sHTMLBody = "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>This is to notify you that an inspection for this permit has been scheduled.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & "<p>Permit #: " & sPermitNo & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Desc: " & sDesc & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Job Site: " & sJobSite & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Inspection: " & sInspectionDescription & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "Scheduled Date: " & Replace(sScheduledDate, "'", "") & "  " & Replace(sScheduledTime, "'", "") & " " & Replace(sScheduledAmPm, "'", "") & "<br />" & vbcrlf
	sHTMLBody = sHTMLBody & "</p>" & vbcrlf & vbcrlf

	If sToUserIdType <> "applicantuserid" Then 
		' no links for the applicant yet
		sHTMLBody = sHTMLBody & "<p><a href=""" & session("egovclientwebsiteurl") & "/admin/permits/permitedit.asp?permitid=" & iPermitId & """ title=""click to view"">Click here to view the details for this permit.</a></p>"
	Else 
		' Get the applicant email address 
		GetPermitApplicantEmailAndName iPermitId, sToEmail, sToName
	End If 

'	response.write sHTMLBody
'	response.write "sToName: " & sToName & "<br />"
'	response.write "sToEmail: " & sToEmail & "<br />"
'	response.write "sFromName: " & sFromName & "<br />"

	If sToEmail <> "" Then 
		SendEmailPermits sToName, sToEmail, sFromName, "noreplies@egovlink.com", sSubject, sHTMLBody
	End If 

End Sub 




'--------------------------------------------------------------------------------------------------
' GetPermitApplicantEmailAndName iPermitId, sToEmail, sToName 
'--------------------------------------------------------------------------------------------------
Sub GetPermitApplicantEmailAndName( ByVal iPermitId, ByRef sApplicantEmail, ByRef sApplicantName )
	Dim sSql, oRs

	sSql = " SELECT ISNULL(firstname,'') AS firstname, ISNULL(lastname,'') AS lastname, "
	sSql = sSQl & " ISNULL(company,'') AS company, ISNULL(email,'') AS email " 
	sSql = sSQl & " FROM egov_permitcontacts WHERE isapplicant = 1 AND ispriorcontact = 0 AND permitid = " & iPermitId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("firstname") <> "" Then 
			sApplicantName = oRs("firstname") & " " & oRs("lastname") & "<br />"
		End If 
		If oRs("company") <> "" And sApplicantName = "" Then 
			sApplicantName = sContact & oRs("company") & "<br />" 
		End If 
		sApplicantEmail = oRs("email")
	Else
		sApplicantName = ""
		sApplicantEmail = ""
	End If 

	oRs.Close
	Set oRs = Nothing 


End Sub  


%>
