<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: criticaldateupdate.asp
' AUTHOR: Steve Loar
' CREATED: 05/21/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This updates expiration dates. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   05/21/2008	Steve Loar - INITIAL VERSION
' 1.1	08/19/2010	Steve Loar - Changed from just expiration dates to critical dates
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sCriticalDate, sSql, iCriticalDateType, sExtra, bUpdate, sDateField, sOriginalCriticalDate
Dim sShowDateType, iPermitStatusId, sActivityComment

iPermitId = CLng(request("permitid"))
'response.write iPermitId & "<br />"
sCriticalDate = dbsafe(request("criticaldate"))
'response.write sCriticalDate & "<br />"
sOriginalCriticalDate = request("originalcriticaldate")
'response.write sOriginalCriticalDate & "<br />"
iCriticalDateType = CLng(request("criticaldatetype"))
'response.write iCriticalDateType & "<br />"

sExtra = ""
bUpdate = True 

Select Case iCriticalDateType
	Case 1
		sDateField = "applieddate"
		sShowDateType = "Applied"
	Case 2
		sDateField = "releaseddate"
		sShowDateType = "Released"
	Case 3
		sDateField = "approveddate"
		sShowDateType = "Approved"
	Case 4
		sDateField = "issueddate"
		sShowDateType = "Issued"
	Case 5
		sDateField = "expirationdate"
		sExtra = ", isexpired = 0, overrideexpiration = 1 "
		sShowDateType = "Expiration"
	Case Else
		bUpdate = False 
End Select 
'response.write bUpdate & "<br />"

If bUpdate Then 
	sSql = "UPDATE egov_permits SET " & sDateField & " = '" & sCriticalDate & "' " & sExtra & " WHERE permitid = " & iPermitId 
	sSql = sSql & " AND orgid = " & session("orgid")
	RunSQL sSql

	iPermitStatusId = GetPermitStatusId( iPermitId )
	sActivityComment = "'" & sShowDateType & " date was changed from " & sOriginalCriticalDate & " to " & sCriticalDate & "'"

	' Put an entry in the log for the change
	MakeAPermitLogEntry iPermitid, "'Date Change'", sActivityComment, "NULL", "NULL", iPermitStatusId, 0, 0, 1, "NULL", "NULL", "NULL", "NULL" 

	response.write "Success: " & sCriticalDate
Else
	response.write "Failed: " & sCriticalDate
End If 


%>
