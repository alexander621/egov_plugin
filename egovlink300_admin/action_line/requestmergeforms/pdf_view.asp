<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: PDF_VIEW.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 02/28/06
' COPYRIGHT: COPYRIGHT 2007 ECLINK, INC.
'			 ALL RIGHTS RESERVED.
'
' DESCRIPTION:  VIEW ATTACHMENT
'
' MODIFICATION HISTORY
' 1.0	02/28/07	JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("action line") = "Y" then
   response.redirect "../../admin/outage_feature_offline.asp"
end if

'CALL SUB TO VIEW SELECTED ATTACHMENT
Call subViewPDF(request("pdfid"))

'------------------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS AND SUBROUTINES
'------------------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' SUB SUBVIEWPDF(IPDFID)
'--------------------------------------------------------------------------------------------------
Sub subViewPDF(ipdfid)

'FIND IN SERVER FILESYSTEM
	sServerPath = server.mappath("../../") & "\custom\pub\" & session("virtualdirectory") & "\pdf_forms"
	sServerPath = sServerPath & "\" & ipdfid & ".pdf" 
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	If oFSO.FileExists(sServerPath)  Then
	 	'FILE FOUND IN SERVER FILESYSTEM
	Else
		 'ERROR FILE NOT FOUND
		  response.write "PDF Not Found!"
		  response.end
	End If
	Set oFSO = Nothing

'OPEN DOCUMENT IN BINARY AND STREAM TO BROWSER
	response.clear
	response.contentType = "application/octet-stream"
	response.addheader "content-disposition", "attachment; filename=form.pdf"
	
	Const adTypeBinary = 1
	Set oStream = Server.CreateObject("ADODB.Stream")
	oStream.Open
	oStream.Type = adTypeBinary
	oStream.LoadFromFile sServerPath
	Response.BinaryWrite oStream.Read
	oStream.Close
	Set objStream = Nothing

End Sub
%>