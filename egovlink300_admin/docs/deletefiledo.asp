<!-- #include file="../includes/common.asp" //-->
<!-- #include file="docscommon.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: deletefiledo.asp
' AUTHOR: Steve Loar
' CREATED: 09/07/2010
' COPYRIGHT: Copyright 2010 eclink, inc.
'			 All Rights Reserved.
'
' Description: Deletes document folders.
'
' MODIFICATION HISTORY
' 1.0   09/07/2010	Steve Loar - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim sSuccessFlag, sTargetFile, iFileId, sPhysicalPath, sFileName, objDir, sDBFile, sSql

sTargetFile = dbsafe(request("path"))

iFileId = CLng(request("fileid"))
sFileName = Trim(request("filename"))

'sPhysicalPath = "e:" & Replace(sTargetFile, "/", "\") ' This is the physical file we need to check and delete
sPhysicalPath = Application("DocumentsDrive") & Replace(sTargetFile, "/", "\") ' This is the physical file we need to check and delete

'sDBFile = Replace(sTargetFile, "egovlink300_docs", "public_documents300")
sDBFile = Replace(sTargetFile, Application("DocumentsRootDirectory"), "public_documents300")

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

If objFSO.FileExists( sPhysicalPath ) Then

	' Delete the file.
	objFSO.DeleteFile(sPhysicalPath)

	sSql = "DELETE FROM Documents WHERE DocumentID = " & iFileId
	'response.write sSql & "<br /><br />"
	RunSQLStatement sSql
	
	sSuccessFlag = "fd"
Else
	sSuccessFlag = "nf"
End If 

Set objFSO = Nothing 

'response.write "sSuccessFlag = " & sSuccessFlag & "<br /><br />"

If sSuccessFlag= "nf" Then
	' file not found so go back to the default page
	response.redirect "default.asp?path=" & request("path") & "&filename=" & sFileName & "&fileid=" & iFileId & "&sf=" & sSuccessFlag
Else
	' file deleted so go to the default page 
	response.redirect "default.asp?filename=" & sFileName & "&sf=" & sSuccessFlag
End If 



%>