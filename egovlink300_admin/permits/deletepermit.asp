<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: deletepermit.asp
' AUTHOR: Steve Loar
' CREATED: 12/11/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This deletes a permit in all related table entries. Called via AJAX
'
' MODIFICATION HISTORY
' 1.0   12/11/2008   Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sSql

iPermitId = CLng(request("permitid"))

sSql = "DELETE FROM egov_permitaddress WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitcontacts_licenses WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitcontacts WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitfeecategories WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitresidentialunitstepfees WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitresidentialunits WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitvaluationstepfees WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitvaluations WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitfixturestepfees WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitfixtures WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitfees WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitinvoiceitems WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitinvoices WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitinspections WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitreviews WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitlog WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permitpermittypes WHERE permitid = " & iPermitId
RunSQL sSql

' Clean out the attachments files
DeleteAttachmentFiles iPermitId

sSql = "DELETE FROM egov_permitattachments WHERE permitid = " & iPermitId
RunSQL sSql

sSql = "DELETE FROM egov_permits WHERE permitid = " & iPermitId
RunSQL sSql

response.write "DELETED"


'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------
Sub DeleteAttachmentFiles( iPermitId ) 
	Dim sSql, oRs, sServerPath, sFilePath, oFSO

	sSql = "SELECT permitattachmentid, fileextension FROM egov_permitattachments "
	sSql = sSql & " WHERE permitid = " & iPermitId

	Set oRs =  Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL,Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		sServerPath = server.mappath("..\permitattachments")
		Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
		Do While Not oRs.EOF
			sFilePath = sServerPath & "\" & oRs("permitattachmentid") & "." & oRs("fileextension")
			If oFSO.FileExists( sFilePath )  Then
				oFSO.DeleteFile( sFilePath )
			End If 
			oRs.MoveNext 
		Loop 
		Set oFSO = Nothing
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


%>
