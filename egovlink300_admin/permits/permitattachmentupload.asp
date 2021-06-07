<!-- #include file="../includes/common.asp" //-->
<!-- #include file="permitcommonfunctions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: permitattachmentupload.asp
' AUTHOR: Steve Loar
' CREATED: 06/19/2008
' COPYRIGHT: Copyright 2008 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Adds fees to permits
'
' MODIFICATION HISTORY
' 1.0   06/19/2008	Steve Loar - INITIAL VERSION
' 2.0	05/03/2010	Steve Loar - Modified to put each city in their own folder
' 2.1	01/11/2011	Steve Loar - Added flag to notify reviewers of attachments
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iPermitId, sAttachmentDesc, oUpload, sFilePath, sFileName, sFileExt, sServerPath, iSuccess

Set oUpload = Server.CreateObject("Dundas.Upload.2")
oUpload.MaxFileSize = (5096000 * 5) ' MAX SIZE OF UPLOAD SPECIFIED IN BYTES, (4096000 * 5) =  APPX. 20MB
oUpload.SaveToMemory

iPermitId = CLng(oUpload.Form("permitid"))

If CLng(oUpload.Files(0).Size) <= CLng(21504000) Then  'APPX. 21MB
	
	sAttachmentDesc = dbsafe(oUpload.Form("attachmentdesc"))

	sFilePath = oUpload.Files(0).OriginalPath
	'	response.write sFilePath & "<br />"
	sFileName = LCase(Right(sFilePath,Len(sFilePath) - InstrRev(sFilePath,"\")))
	'	response.write sFileName & "<br />"
	sFileExt = LCase(Right(sFileName,Len(sFileName) - InstrRev(sFileName,".")))
	'	response.write sFileExt & "<br />"

	sServerPath = server.mappath("..\permitattachments\" & session("sitename"))
	'	response.write sServerPath & "<br />"

	sServerFileName = StoreAttatchmentInfo( iPermitId, sFileName, sAttachmentDesc, sFileExt, "..\permitattachments\" & session("sitename") )
	sServerFileName = sServerFileName '& "." & sFileExt 

	oUpload.Files(0).SaveAs(  sServerPath & "\" & sServerFileName )
	iSuccess = 1
Else
	iSuccess = 0
End If 

Set oUpload = Nothing
'response.write "<br /><br />Upload complete."
If iSuccess = 1 Then 
	If ReviewersNeedNotification( iPermitId ) Then
		NotifyReviewersOfAttachment iPermitId
	End If 
End If

response.redirect "permitattachment.asp?permitid=" & iPermitId & "&success=" & iSuccess

%>
<html>
	<head>
		<script language="Javascript">
		<!--

		function doClose()
		{
			window.close();
			window.opener.focus();
		}

		//-->
		</script>

	</head>
	<body onload="doClose();">
		<p>Upload complete.</p>
	</body>
</html>

<%

'--------------------------------------------------------------------------------------------------
' string StoreAttatchmentInfo( iPermitId, sFileName, sAttachmentDesc, sFileExt, sAttachmentPath )
'--------------------------------------------------------------------------------------------------
Function StoreAttatchmentInfo( ByVal iPermitId, ByVal sFileName, ByVal sAttachmentDesc, ByVal sFileExt, ByVal sAttachmentPath )
	Dim sSql, iPermitAttachmentId

	sSql = "INSERT INTO egov_permitattachments ( permitid, orgid, attachmentname, fileextension, "
	sSql = sSql & " description, adminuserid, dateadded, attachmentpath ) VALUES ( "
	sSql = sSql & iPermitid & ", " & session("OrgID") & ", '" & dbsafe(sFileName) & "', '" & sFileExt & "', '" & dbsafe(sAttachmentDesc) 
	sSql = sSql & "', " & session("UserID") & ", dbo.GetLocalDate(" & Session("OrgID") & ",getdate()), '" & sAttachmentPath & "' )"
	session("StoreAttatchmentInfo") = sSql

	iPermitAttachmentId = RunIdentityInsert( sSql )

'	sSql = "UPDATE egov_permitattachments SET attachmentname = '" & iPermitAttachmentId & "_" & sFileName
'	sSql = sSql & "' WHERE permitattachmentid = " & iPermitAttachmentId
'	RunSQL sSql
	session("StoreAttatchmentInfo") = ""
	
	StoreAttatchmentInfo = iPermitAttachmentId & "_" & sFileName

End Function 


'--------------------------------------------------------------------------------------------------
' boolean ReviewersNeedNotification( iPermitId )
'--------------------------------------------------------------------------------------------------
Function ReviewersNeedNotification( ByVal iPermitId )
	Dim sSql, oRs

	sSql = "SELECT attachmentrevieweralert FROM egov_permitpermittypes "
	sSql = sSql & " WHERE permitid = '" & iPermitId & "'"
	session("ReviewersNeedNotification") = sSql

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then
		If oRs("attachmentrevieweralert") Then 
			ReviewersNeedNotification = True 
		Else
			ReviewersNeedNotification = False 
		End If 
	Else
		ReviewersNeedNotification = False 
	End If
	
	oRs.Close
	Set oRs = Nothing 
	
	session("ReviewersNeedNotification") = ""

End Function 


'--------------------------------------------------------------------------------------------------
' void NotifyReviewersOfAttachment iPermitId 
'--------------------------------------------------------------------------------------------------
Sub NotifyReviewersOfAttachment( ByVal iPermitId )
	Dim sSql, oRs, sToName, sToEmail, sFromName, sFromEmail, sSubject, sHTMLBody
	Dim sPermitNo, sDesc, sJobSite, iPermitTypeId, sOrgName, sStatus, sLocation
	Dim sLocationType

	' Pull the permit details needed
	sPermitNo = GetPermitNumber( iPermitId ) '	in permitcommonfunctions.asp
	sDesc = GetPermitTypeDesc( iPermitId, True ) '	in permitcommonfunctions.asp
	sJobSite = GetPermitJobSite( iPermitId ) '	in permitcommonfunctions.asp
	iPermitTypeId = GetPermitTypeId( iPermitId ) '	in permitcommonfunctions.asp
	sStatus = GetPermitStatusByPermitId( iPermitId ) '	in permitcommonfunctions.asp
	sLocation = Replace(GetPermitPermitLocation( iPermitId ), Chr(10), Chr(10) & "<br />")
	sLocationType = GetPermitLocationType( iPermitId )	'	in permitcommonfunctions.asp

	sSubject = "Permit " & sPermitNo & " - New Attachment" 
	sOrgName = GetOrgName( session("orgid") )
	sFromName = sOrgName & " E-GOV WEBSITE"

	' Build the email body
	sHTMLBody = "<p>This automated message was sent by the " & sOrgName & " E-Gov web site. Do not reply to this message.</p>" & vbcrlf  & vbcrlf 
	sHTMLBody = sHTMLBody & vbcrlf & vbcrlf & "<p>This permit has a new attachment.</p>" & vbcrlf  & vbcrlf 
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
	sSql = sSql & " WHERE U.userid = R.revieweruserid AND U.isdeleted = 0 AND R.permitid = " & iPermitId

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



%>
